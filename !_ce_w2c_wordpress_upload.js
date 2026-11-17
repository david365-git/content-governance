/* ============================================================
   W2C — WORDPRESS MEDIA UPLOAD
   Dedicated W2C upload endpoint. Does not create or publish posts.
   Required Apps Script Script Properties:
   W2C_WP_SITE_URL  e.g. https://www.abbeyfloorcare.co.uk
   W2C_WP_USERNAME  WordPress username with upload_files permission
   W2C_WP_APP_PASSWORD  WordPress Application Password
============================================================ */

function w2cUploadImageToWordPress(payload) {
  if (!payload || typeof payload !== 'object') {
    return { success: false, message: 'Missing image upload payload.' };
  }
  var props = PropertiesService.getScriptProperties();
  var site = String(props.getProperty('W2C_WP_SITE_URL') || '').trim().replace(/\\/+$/, '');
  var user = String(props.getProperty('W2C_WP_USERNAME') || '').trim();
  var pass = String(props.getProperty('W2C_WP_APP_PASSWORD') || '').trim();
  if (!/^https:\/\/[^\\s/]+(?:\/[^?#]*)?$/i.test(site) || !user || !pass) {
    return { success: false, message: 'WordPress upload is not configured. Set W2C_WP_SITE_URL, W2C_WP_USERNAME and W2C_WP_APP_PASSWORD in Apps Script Script Properties.' };
  }

  var name = String(payload.name || '').replace(/^.*[\\/]/, '').trim();
  var match = name.match(/\\.(webp|png|jpe?g|gif|avif)$/i);
  if (!match) return { success: false, message: 'Unsupported image type: ' + name };
  var ext = match[1].toLowerCase();
  var mime = ({webp:'image/webp',png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',avif:'image/avif'})[ext];
  var raw = String(payload.base64 || '');
  if (!raw || raw.length > 12 * 1024 * 1024) {
    return { success: false, message: 'Image is empty or exceeds the 8 MB upload limit.' };
  }
  var bytes;
  try { bytes = Utilities.base64Decode(raw); }
  catch (e) { return { success: false, message: 'Invalid image data.' }; }
  if (!bytes.length || bytes.length > 8 * 1024 * 1024) {
    return { success: false, message: 'Image must be between 1 byte and 8 MB.' };
  }

  var auth = 'Basic ' + Utilities.base64Encode(user + ':' + pass);
  var safeName = name.replace(/[^a-zA-Z0-9._-]/g, '-');
  var endpoint = site + '/wp-json/wp/v2/media';
  var result;
  try {
    result = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: mime,
      headers: {
        Authorization: auth,
        'Content-Disposition': 'attachment; filename="' + safeName + '"'
      },
      payload: bytes,
      muteHttpExceptions: true,
      followRedirects: false
    });
  } catch(e) {
    return { success: false, message: 'WordPress media request failed: ' + e.message };
  }

  var code = result.getResponseCode();
  var data;
  try { data = JSON.parse(result.getContentText()); } catch(e) { data = {}; }
  if (code < 200 || code >= 300 || !data.id || !data.source_url) {
    var wpError = String(data.message || 'HTTP ' + code);
    return { success: false, message: 'WordPress rejected ' + safeName + ': ' + wpError.slice(0, 240) };
  }

  var url = String(data.source_url);
  // Prevent inserting a non-WordPress or unrelated image URL.
  var expectedHost = site.match(/^https:\/\/([^/]+)/i)[1].toLowerCase();
  var actualHost = (url.match(/^https:\/\/([^/]+)/i) || [,''])[1].toLowerCase();
  if (actualHost !== expectedHost && actualHost !== expectedHost.replace(/^www\\./, '') &&
      actualHost.replace(/^www\\./, '') !== expectedHost.replace(/^www\\./, '')) {
    return { success: false, message: 'WordPress returned an unexpected media host. Media ID: ' + data.id + '. Check the WordPress media settings before inserting the URL.' };
  }

  var metaStatus = '';
  var alt = String(payload.alt || '').slice(0, 500);
  var caption = String(payload.caption || '').slice(0, 1200);
  if (alt || caption) {
    try {
      var metaResponse = UrlFetchApp.fetch(endpoint + '/' + encodeURIComponent(data.id), {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: auth },
        payload: JSON.stringify({alt_text: alt, caption: caption}),
        muteHttpExceptions: true,
        followRedirects: false
      });
      if (metaResponse.getResponseCode() < 200 || metaResponse.getResponseCode() >= 300) {
        metaStatus = 'Uploaded, but WordPress did not save alt text/caption.';
      }
    } catch(e) {
      metaStatus = 'Uploaded, but alt text/caption update failed.';
    }
  }
  return { success: true, id: data.id, url: url, message: metaStatus || 'Uploaded to WordPress Media Library.' };
}
