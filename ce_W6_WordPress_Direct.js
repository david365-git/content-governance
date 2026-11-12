/**
 * W6 direct WordPress publishing — server-side Apps Script.
 * Requires the Abbey W6 REST bridge on WordPress (not provided by core WP REST).
 * Does not use the WPVibe connector. Never expose credentials to the sidebar.
 */
function ce_w6WpSettings_() {
  var p = PropertiesService.getScriptProperties();
  var site = String(p.getProperty('CE_WP_MEDIA_SITE_URL') || '').trim().replace(/\/+$/, '');
  var username = String(p.getProperty('CE_WP_MEDIA_USERNAME') || '').trim();
  var password = String(p.getProperty('CE_WP_MEDIA_APP_PASSWORD') || '').trim();
  if (!/^https:\/\/[^/]+$/i.test(site) || !username || !password) {
    throw new Error('WordPress Application Password settings are not configured.');
  }
  return {
    site: site,
    authorization: 'Basic ' + Utilities.base64Encode(username + ':' + password)
  };
}

function ce_w6WpRequest_(method, path, body) {
  var settings = ce_w6WpSettings_();
  var options = {
    method: method,
    headers: {Authorization: settings.authorization, Accept: 'application/json'},
    muteHttpExceptions: true
  };
  if (body !== undefined) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }
  var response = UrlFetchApp.fetch(settings.site + path, options);
  var status = response.getResponseCode();
  var parsed = {};
  try { parsed = JSON.parse(response.getContentText()); } catch (ignore) {}
  if (status < 200 || status >= 300) {
    throw new Error('WordPress HTTP ' + status + ': ' +
      String(parsed.message || response.getContentText()).slice(0, 350));
  }
  return parsed;
}

function ce_w6CurrentRow_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  if (!posts) throw new Error('posts sheet not found.');
  var row = posts.getActiveRange().getRow();
  if (row < 2) throw new Error('Select a post row first.');
  var headers = posts.getRange(1, 1, 1, posts.getLastColumn()).getDisplayValues()[0];
  var ix = headers.indexOf('Post ID');
  if (ix < 0) throw new Error('Post ID column not found.');
  var postId = Number(posts.getRange(row, ix+1).getValue());
  if (!Number.isSafeInteger(postId) || postId <= 0) throw new Error('Invalid Post ID.');
  return postId;
}

/** Read-only preflight. Confirms the live post and current published dates. */
function ce_w6WordPressPreflight(request) {
  try {
    var postId = ce_w6CurrentRow_();
    if (!request || Number(request.postId) !== postId) {
      throw new Error('Selected post changed. Reload W6 before proceeding.');
    }
    var post = ce_w6WpRequest_('get',
      '/wp-json/wp/v2/posts/' + postId + '?context=edit&_fields=id,status,link,date,date_gmt,modified,modified_gmt,title');
    if (Number(post.id) !== postId || post.status !== 'publish') {
      throw new Error('WordPress did not confirm an existing published post.');
    }
    return {success:true, postId:postId, title:post.title && post.title.raw || '',
      url:post.link, publishedDate:post.date, modifiedDate:post.modified,
      message:'WordPress post verified. No changes made.'};
  } catch (e) {
    return {success:false,message:String(e.message || e)};
  }
}

/**
 * Publish only after the matching row was loaded, and the user explicitly
 * confirmed UPDATE or REPUBLISH. Bridge validates WordPress write atomically.
 * IMPORTANT: REPUBLISH date resets require syndication verification before
 * production use. This implementation blocks republishing until enabled.
 */
function ce_w6WordPressPublish(request) {
  try {
    if (!request || !['update','republish'].includes(request.mode)) {
      throw new Error('Choose Update or Republish.');
    }
    var postId = ce_w6CurrentRow_();
    if (Number(request.postId) !== postId) {
      throw new Error('Selected post changed. Reload W6.');
    }
    if (request.mode === 'republish') {
      throw new Error('Republish is disabled until Fabritec syndication trigger is verified.');
    }
    var data = ce_getW6Data();
    if (!data.success || Number(data.postId) !== postId) {
      throw new Error('W6 row data could not be verified.');
    }
    if (!String(data.newH1 || '').trim() || !String(data.newHtml || '').trim()) {
      throw new Error('New H1 and New HTML are required.');
    }
    if (!String(data.newSchema || '').trim()) {
      throw new Error('New Schema is required; the existing schema will not be overwritten with blank.');
    }

    // Preflight is repeated server-side immediately before the write.
    var current = ce_w6WordPressPreflight({postId:postId});
    if (!current.success) throw new Error(current.message);

    // Optional EM image: only send the URL when populated.
    // The WordPress bridge resolves and validates the Media Library attachment.
    var imageUrl = String(data.featuredImageUrl || '').trim();
    if (imageUrl && !/^https:\/\/[^\\s?#]+(?:\\?[^#\\s]*)?$/i.test(imageUrl)) {
      throw new Error('Column EM must contain a valid HTTPS featured-image URL.');
    }
    var publishBody = {
      post_id: postId,
      mode: 'update',
      expected_date: current.publishedDate,
      title: data.newH1,
      html: data.newHtml,
      meta_title: data.newMetaTitle,
      meta_description: data.newMetaDesc,
      focus_keyphrase: data.yoastKeyphrase,
      schema: data.newSchema
    };
    if (imageUrl) publishBody.featured_image_url = imageUrl;
    var result = ce_w6WpRequest_('post', '/wp-json/abbey/v1/w6-publish', publishBody);
    if (!result.success || Number(result.post_id) !== postId ||
        !result.published_date || !result.modified_date ||
        result.published_date !== current.publishedDate) {
      throw new Error('WordPress response was not fully verified. Check the post before retrying.');
    }
    return {
      success:true, postId:postId, mode:'update',
      publishedDate:result.published_date,
      modifiedDate:result.modified_date,
      url:result.url || current.url,
      featuredImageUrl: result.featured_image_url || '',
      featuredImageAction: result.featured_image_action || 'unchanged',
      message:'WordPress updated. Published Date retained. Featured image: ' +
        (result.featured_image_action || 'unchanged') + '. Verify page and then update the log.'
    };
  } catch (e) {
    return {success:false,message:String(e.message || e)};
  }
}
