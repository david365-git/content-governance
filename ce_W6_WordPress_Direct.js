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
 * Republish uses a separate guarded WordPress operation. Verify the external
 * syndication event independently; a WordPress date change alone is not proof.
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
    if (request.mode === 'republish' &&
        (!request.verifiedPublishedDate || !request.verifiedModifiedDate)) {
      throw new Error('Verify the live publication and modification dates before republishing.');
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
    if (request.mode === 'republish' &&
        (current.publishedDate !== request.verifiedPublishedDate ||
         current.modifiedDate !== request.verifiedModifiedDate)) {
      throw new Error('WordPress post changed since the confirmation. Verify again before republishing.');
    }

    // Optional EM image: only send the URL when populated.
    // The WordPress bridge resolves and validates the Media Library attachment.
    var imageUrl = String(data.featuredImageUrl || '').trim();
    if (imageUrl && !/^https:\/\/[^\s?#]+(?:\?[^#\s]*)?$/i.test(imageUrl)) {
      throw new Error('Column EM must contain a valid HTTPS featured-image URL.');
    }
    // Fail closed if a saved featured image has not yet been reflected in W5C.
    if (imageUrl) {
      var rawSchema = String(data.newSchema || '').trim();
      var bodyMatch = rawSchema.match(/<script\b[^>]*>([\s\S]*?)<\/script>/i);
      var parsedSchema;
      try { parsedSchema = JSON.parse(bodyMatch ? bodyMatch[1] : rawSchema); }
      catch (schemaError) { throw new Error('W5C schema is invalid; rebuild before publishing.'); }
      var nodes = Array.isArray(parsedSchema['@graph']) ? parsedSchema['@graph'] : [parsedSchema];
      var matchingPrimaryImage = nodes.some(function(node) {
        if (!node || !node['@type']) return false;
        var types = Array.isArray(node['@type']) ? node['@type'] : [node['@type']];
        if (!types.some(function(t) { return t === 'Article' || t === 'Service'; })) return false;
        var images = Array.isArray(node.image) ? node.image : [node.image];
        return images.some(function(img) {
          return img === imageUrl || (img && typeof img === 'object' &&
            (img.url === imageUrl || img.contentUrl === imageUrl));
        });
      });
      if (!matchingPrimaryImage) {
        throw new Error('Column EM and W5C schema image differ. Refresh W5C before publishing.');
      }
    }
    var publishBody = {
      post_id: postId,
      mode: request.mode,
      expected_date: current.publishedDate,
      title: data.newH1,
      html: data.newHtml,
      meta_title: data.newMetaTitle,
      meta_description: data.newMetaDesc,
      focus_keyphrase: data.yoastKeyphrase,
      schema: data.newSchema
    };
    if (request.mode === 'republish') publishBody.expected_modified = current.modifiedDate;
    if (imageUrl) publishBody.featured_image_url = imageUrl;
    var result = ce_w6WpRequest_('post', '/wp-json/abbey/v1/w6-publish', publishBody);
    if (!result.success || Number(result.post_id) !== postId ||
        !result.published_date || !result.modified_date ||
        !result.schema ||
        (request.mode === 'update'
          ? result.published_date !== current.publishedDate
          : result.published_date === current.publishedDate)) {
      throw new Error('WordPress response was not fully verified. Check the post before retrying.');
    }
    return {
      success:true, postId:postId, mode:request.mode,
      publishedDate:result.published_date,
      modifiedDate:result.modified_date,
      schema:result.schema,
      url:result.url || current.url,
      featuredImageUrl: result.featured_image_url || '',
      featuredImageAction: result.featured_image_action || 'unchanged',
      message:'WordPress ' + (request.mode === 'republish' ? 'republished' : 'updated') +
        '. ' + (request.mode === 'republish' ? 'Published Date advanced' : 'Published Date retained') +
        '. Featured image: ' + (result.featured_image_action || 'unchanged') +
        '. Syndication delivery requires separate verification.'
    };
  } catch (e) {
    return {success:false,message:String(e.message || e)};
  }
}



/**
 * Keep the posts source Schema (JSON-LD) consistent with the canonical
 * schema returned by WordPress, without running any WordPress write.
 * Any failure is reported as a partial recording issue, not a publish failure.
 */
function ce_w6SaveConfirmedSchema_(postId, expectedSchema, confirmedSchema) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
    if (!sheet) throw new Error('posts sheet missing.');
    var row = sheet.getActiveRange().getRow();
    if (row < 2) throw new Error('Select a posts data row.');
    var h = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
    var idCol = h.indexOf('Post ID') + 1;
    var schemaCol = h.indexOf('Schema (JSON-LD)') + 1;
    if (!idCol || !schemaCol) throw new Error('Post ID or Schema (JSON-LD) column missing.');
    if (String(sheet.getRange(row,idCol).getValue()).trim() !== String(postId)) {
      throw new Error('Active posts row changed; schema not overwritten.');
    }
    var cell = sheet.getRange(row,schemaCol);
    var existing = String(cell.getValue() || '').trim();
    if (existing !== String(expectedSchema || '').trim() &&
        existing !== String(confirmedSchema || '').trim()) {
      throw new Error('Schema source changed since publishing; manual reconciliation required.');
    }
    cell.setValue(String(confirmedSchema || ''));
    SpreadsheetApp.flush();
    return {success:true};
  } catch (e) {
    return {success:false,message:String(e.message || e)};
  }
}

/**
 * Single W6 action: publish first, then record the confirmed WordPress result.
 * If a sheet write fails after publishing, report a PARTIAL result and do not
 * invite another publish; the user can reconcile the record separately.
 */
function ce_w6DirectPublishAndRecord(request) {
  var original = ce_getW6Data();
  if (!original.success || !request ||
      String(original.postId) !== String(request.postId)) {
    return {success:false,published:false,message:'The selected post changed. Reload W6.'};
  }
  var result = ce_w6WordPressPublish(request);
  if (!result.success) {
    return {success:false,published:false,message:result.message};
  }
  var problems = [];
  var sourceSchema = original.newSchema;
  original.newPublishedDate = result.publishedDate;
  original.newUpdatedDate = result.modifiedDate;
  // Log the canonical JSON-LD returned by WordPress after it has saved the post.
  original.newSchema = result.schema;
  original.publicationAction = result.mode;
  original.url = result.url || original.url;
  var log = ce_updatePublicationLog(original);
  if (!log.success) problems.push('Publication log: ' + log.message);
  var exp = ce_updateSiteExport(result);
  if (!exp.success) problems.push('site-export: ' + exp.message);
  result.published = true;
  result.logUpdated = !!log.success;
  result.siteExportUpdated = !!exp.success;
  var schemaSync = ce_w6SaveConfirmedSchema_(result.postId, sourceSchema, result.schema);
  if (!schemaSync.success) problems.push('posts schema: ' + schemaSync.message);
  result.postsSchemaUpdated = !!schemaSync.success;
  if (problems.length) {
    result.success = false;
    result.message = 'WordPress WAS updated, but sheet recording needs attention. Do not publish again. ' +
      problems.join(' | ');
  } else {
    result.message = 'WordPress ' + (result.mode === 'republish' ? 'republished' : 'updated') +
      ' and verified; publication log and site-export recorded. ' +
      (result.mode === 'republish' ? 'Published date advanced.' : 'Published date preserved.') +
      ' Featured image: ' + result.featuredImageAction +
      '. Verify syndication separately if republished.';
  }
  return result;
}
