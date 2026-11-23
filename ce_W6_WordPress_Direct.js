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
    // A server may commit wp_update_post() and then return an empty 500 when
    // a later hook/meta operation dies. Preserve the HTTP status and WP code.
    var raw = String(response.getContentText() || '').trim();
    var detail = parsed && (parsed.message || parsed.code) ||
      (raw ? raw.slice(0, 350) : 'empty response body; server-side error details unavailable');
    var code = parsed && parsed.code ? ' [' + String(parsed.code).slice(0, 90) + ']' : '';
    throw new Error('WordPress HTTP ' + status + code + ': ' + String(detail).slice(0, 350));
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
    if (Number(post.id) !== postId || !['publish','future'].includes(post.status)) {
      throw new Error('WordPress did not confirm an existing published or scheduled post.');
    }
    return {success:true, postId:postId, title:post.title && post.title.raw || '',
      url:post.link, status:post.status, publishedDate:post.date, modifiedDate:post.modified,
      message:'WordPress ' + (post.status === 'future' ? 'scheduled' : 'published') + ' post verified. No changes made.'};
  } catch (e) {
    return {success:false,message:String(e.message || e)};
  }
}


/**
 * Recovery after an uncertain WordPress write. Standard read-only REST only:
 * no additional publish, no sheet write, and no claim that schema was saved.
 */
function ce_w6ReadBackAfterError_(postId, mode, before, intended, originalError) {
  var reason = String(originalError && (originalError.message || originalError) || 'Unknown WordPress error');
  var base = {
    success:false, postId:postId, mode:mode, published:false,
    partial:true, retryBlocked:true, recoveryCheck:'unavailable',
    checks:{}, diagnostics:[],
    message:'W6 received no verified WordPress receipt. A post write may have occurred. ' +
      'DO NOT RETRY publishing. Original error: ' + reason
  };
  try {
    // Standard authenticated REST read-back only. Never issue another POST.
    var live = ce_w6WpRequest_('get',
      '/wp-json/wp/v2/posts/' + postId +
      '?context=edit&_fields=id,status,date,modified,link,title,content,meta,featured_media');
    if (Number(live.id) !== postId) throw new Error('WordPress returned a different post ID.');
    var liveDate = String(live.date || '');
    var liveStatus = String(live.status || '');
    var changed = liveDate !== String(before.publishedDate || '') ||
      liveStatus !== String(before.status || '');
    var title = live.title && live.title.raw;
    var html = live.content && live.content.raw;
    var meta = live.meta || {};
    var checks = {
      status:liveStatus, publicationDate:liveDate,
      titleMatches:typeof title === 'string' ? title === String(intended.newH1 || '') : null,
      htmlMatches:typeof html === 'string' ? html === String(intended.newHtml || '') : null,
      metaTitleMatches:typeof meta._yoast_wpseo_title === 'string'
        ? meta._yoast_wpseo_title === String(intended.newMetaTitle || '') : null,
      metaDescriptionMatches:typeof meta._yoast_wpseo_metadesc === 'string'
        ? meta._yoast_wpseo_metadesc === String(intended.newMetaDesc || '') : null,
      focusKeyphraseMatches:typeof meta._yoast_wpseo_focuskw === 'string'
        ? meta._yoast_wpseo_focuskw === String(intended.yoastKeyphrase || '') : null,
      featuredMediaId:Number(live.featured_media || 0),
      featuredImageMatches:null,
      schemaSaved:null
    };
    if (String(intended.featuredImageUrl || '').trim()) {
      // Read attachment metadata to compare against EM without changing it.
      if (checks.featuredMediaId > 0) {
        try {
          var media = ce_w6WpRequest_('get', '/wp-json/wp/v2/media/' +
            checks.featuredMediaId + '?context=edit&_fields=id,source_url');
          checks.featuredImageMatches = String(media.source_url || '').replace(/^http:/i,'https:') ===
            String(intended.featuredImageUrl || '').trim().replace(/^http:/i,'https:');
        } catch (mediaError) {
          base.diagnostics.push('Featured image URL could not be read: ' +
            String(mediaError.message || mediaError).slice(0,200));
        }
      } else {
        checks.featuredImageMatches = false;
      }
    }
    // Head & Footer Code is NOT registered in core WP REST post meta.
    // Therefore schemaSaved is UNKNOWN, not true or false, until an
    // independently authorised schema/head-code read verifies it.
    base.publishedDate = liveDate;
    base.modifiedDate = String(live.modified || '');
    base.status = liveStatus;
    base.url = String(live.link || '');
    base.checks = checks;
    base.recoveryCheck = 'read_only_complete';
    base.published = mode === 'republish' && liveStatus === 'publish' && changed;
    var mismatch = Object.keys(checks).filter(function(k) {
      return /Matches$/.test(k) && checks[k] === false;
    });
    var unknown = Object.keys(checks).filter(function(k) {
      return (/Matches$/.test(k) || k === 'schemaSaved') && checks[k] === null;
    });
    base.mismatches = mismatch;
    base.unverified = unknown;
    base.message = (base.published
      ? 'WORDPRESS IS ALREADY PUBLISHED (Post ' + postId + ', date ' + liveDate + '). '
      : changed
        ? 'WORDPRESS CHANGED AFTER W6 ERROR (Post ' + postId + ', status ' + liveStatus + '). '
        : 'WordPress status/date were unchanged after W6 error (Post ' + postId + '). ') +
      'The request has NO VERIFIED RECEIPT; post, metadata and schema changes may be partial. ' +
      (mismatch.length ? 'Differences detected: ' + mismatch.join(', ') + '. ' : '') +
      (unknown.length ? 'Not independently verified: ' + unknown.join(', ') + '. ' : '') +
      'No publication-log, site-export or posts schema records were written by W6. ' +
      'DO NOT REPUBLISH OR RETRY UPDATE. Reconcile individual fields without publishing. ' +
      'Original error: ' + reason;
  } catch (readError) {
    base.message += ' Read-only recovery failed: ' + String(readError.message || readError) +
      '. No safe assumptions can be made.';
  }
  return base;
}

/**
 * Publish only after the matching row was loaded, and the user explicitly
 * confirmed UPDATE or REPUBLISH. WordPress writes are NOT atomic; the
 * bridge response may fail after the post has changed.
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
    if (!['publish','future'].includes(current.status)) throw new Error('WordPress post status is not eligible.');
    if (request.verifiedStatus && request.verifiedStatus !== current.status) {
      throw new Error('WordPress post status changed since verification. Verify again.');
    }
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
      expected_status: current.status,
      title: data.newH1,
      html: data.newHtml,
      meta_title: data.newMetaTitle,
      meta_description: data.newMetaDesc,
      focus_keyphrase: data.yoastKeyphrase,
      schema: data.newSchema
    };
    if (request.mode === 'republish') publishBody.expected_modified = current.modifiedDate;
    if (imageUrl) publishBody.featured_image_url = imageUrl;
    // After submitting a write, any transport, HTTP or receipt-validation
    // failure is potentially a PARTIAL write. Read back; never retry automatically.
    var result;
    try {
      result = ce_w6WpRequest_('post', '/wp-json/abbey/v1/w6-publish', publishBody);
      if (!result.success || Number(result.post_id) !== postId ||
          !result.published_date || !result.modified_date ||
          !result.status ||
          (request.mode === 'update' && result.status !== current.status) ||
          (request.mode === 'republish' && result.status !== 'publish') ||
          !result.schema ||
          (request.mode === 'update'
            ? result.published_date !== current.publishedDate
            : result.published_date === current.publishedDate)) {
        throw new Error('WordPress response did not verify all required fields.');
      }
    } catch (writeError) {
      return ce_w6ReadBackAfterError_(postId, request.mode, current, data, writeError);
    }
    return {
      success:true, postId:postId, mode:request.mode,
      publishedDate:result.published_date,
      modifiedDate:result.modified_date,
      previousStatus:current.status,
      status:result.status,
      schema:result.schema,
      url:result.url || current.url,
      featuredImageUrl: result.featured_image_url || '',
      featuredImageAction: result.featured_image_action || 'unchanged',
      message:'WordPress ' + (request.mode === 'republish' ? (current.status === 'future' ? 'scheduled post published now' : 'republished') : 'updated') +
        '. ' + (request.mode === 'republish' ? 'Published date set to now' : 'Publication/scheduled date and status retained') +
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
 * Record the final successful W6 action by Post ID, never by a potentially
 * changed active row. The posts!D140 colour is #ffff00 (yellow/black).
 * Queue E is its existing Status field; prior workflow stamps remain intact.
 * This is ONLY called after WordPress and the normal W6 sheet writes succeed.
 */
function ce_w6MarkCompletedAction_(postId, mode) {
  if (!/^\d+$/.test(String(postId)) || !['republish','update'].includes(mode))
    return {success:false,message:'Invalid W6 completion status.'};
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var posts=ss.getSheetByName('posts');
  if (!posts) return {success:false,message:'posts sheet missing.'};
  var last=posts.getLastRow();
  if(last<2)return {success:false,message:'posts sheet has no Post IDs.'};
  var ids=posts.getRange(2,4,last-1,1).getDisplayValues();
  var matches=[];
  ids.forEach(function(r,i){if(String(r[0]).trim()===String(postId))matches.push(i+2);});
  if(matches.length!==1)return {success:false,message:'posts Post ID '+postId+
    ' matched '+matches.length+' rows; completion colour not recorded.'};

  var queue=ss.getSheetByName('Google Love Batch Queue');
  var queueRow=-1;
  if(queue&&queue.getLastRow()>1){
    var qids=queue.getRange(2,1,queue.getLastRow()-1,1).getDisplayValues();
    var qMatches=[];
    qids.forEach(function(r,i){if(String(r[0]).trim()===String(postId))qMatches.push(i+2);});
    if(qMatches.length>1)return {success:false,message:'Duplicate Post ID '+postId+
      ' in Google Love Batch Queue; no completion colours written.'};
    if(qMatches.length===1)queueRow=qMatches[0];
  }
  var colour=mode==='republish'?'#ffff00':'#7030a0';
  var textColour=mode==='republish'?'#000000':'#ffffff';
  try {
    posts.getRange(matches[0],4).setBackground(colour).setFontColor(textColour);
    if(queueRow>1){
      queue.getRange(queueRow,1).setBackground(colour).setFontColor(textColour);
      queue.getRange(queueRow,5).setValue(mode==='republish'?'Republish':'Updated');
    }
    SpreadsheetApp.flush();
    return {success:true,postsRow:matches[0],queueRow:queueRow,
      message:queueRow>1?'Posts and Google Love Batch Queue updated.':
        'Posts coloured; Post ID not present in Google Love Batch Queue.'};
  } catch(e) {
    return {success:false,message:'W6 completion marking failed: '+String(e.message||e)};
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
    // Preserve recovery evidence and do not attempt publication-log writes
    // without the bridge's canonical schema and confirmed success receipt.
    return Object.assign({success:false}, result);
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
  if (!problems.length) {
    // No marking on an unverified or partially recorded WordPress operation.
    var marked = ce_w6MarkCompletedAction_(result.postId, result.mode);
    result.completionMarked = !!marked.success;
    result.completionMarkMessage = marked.message;
    if (!marked.success) problems.push('Completion colours/status: ' + marked.message);
  }
  if (problems.length) {
    result.success = false;
    result.message = 'WordPress WAS updated, but sheet recording needs attention. Do not publish again. ' +
      problems.join(' | ');
  } else {
    result.message = 'WordPress ' + (result.mode === 'republish' ? (result.previousStatus === 'future' ? 'scheduled post published now' : 'republished') : 'updated') +
      ' and verified; publication log and site-export recorded. ' +
      (result.mode === 'republish' ? 'Published date set to now.' : 'Publication/scheduled date and status preserved.') +
      ' Featured image: ' + result.featuredImageAction +
      '. Verify syndication separately if republished.';
  }
  return result;
}
