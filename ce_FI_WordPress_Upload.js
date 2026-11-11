/**
 * FI media upload. Configure the following SCRIPT properties in Apps Script:
 * CE_WP_MEDIA_SITE_URL = https://www.abbeyfloorcare.co.uk (or the target site)
 * CE_WP_MEDIA_USERNAME = WordPress user with upload_files permission
 * CE_WP_MEDIA_APP_PASSWORD = WordPress Application Password
 * Never put the password into source code or the sidebar.
 */
function ceFiUploadImageToWordPress(request) {
  var uploadedUrl = "";
  try {
    if (!request || !request.postId || !request.base64) {
      return {success:false,message:"Generate the image prompt and select an image first."};
    }
    var filename = String(request.filename || "").trim();
    var altText = String(request.altText || "").trim();
    var postId = String(request.postId).trim();
    if (!/^[a-z0-9][a-z0-9-]*\.webp$/i.test(filename)) {
      return {success:false,message:"A valid governed .webp filename is required."};
    }
    if (request.mimeType !== "image/webp") {
      return {success:false,message:"Convert the selected image to WebP before uploading."};
    }
    var settings = PropertiesService.getScriptProperties();
    var site = String(settings.getProperty("CE_WP_MEDIA_SITE_URL") || "").trim().replace(/\/+$/, "");
    var username = String(settings.getProperty("CE_WP_MEDIA_USERNAME") || "").trim();
    var password = String(settings.getProperty("CE_WP_MEDIA_APP_PASSWORD") || "").trim();
    if (!/^https:\/\/[^/]+$/i.test(site) || !username || !password) {
      return {success:false,message:"Configure CE_WP_MEDIA_SITE_URL, CE_WP_MEDIA_USERNAME and CE_WP_MEDIA_APP_PASSWORD in Apps Script Script Properties first."};
    }
    var data = String(request.base64 || "");
    if (data.length > 3*1024*1024 || !/^[a-zA-Z0-9+/=]+$/.test(data)) {
      return {success:false,message:"Invalid or oversized image data (2 MB limit)."};
    }
    var bytes = Utilities.base64Decode(data);
    if (!bytes.length || bytes.length > 2*1024*1024) {
      return {success:false,message:"Image must be smaller than 2 MB."};
    }
    // Verify RIFF/WEBP signature, not just the browser-supplied MIME type.
    var first = bytes.slice(0,12).map(function(x){return x < 0 ? x+256 : x;});
    if (String.fromCharCode.apply(null, first.slice(0,4)) !== "RIFF" ||
        String.fromCharCode.apply(null, first.slice(8,12)) !== "WEBP") {
      return {success:false,message:"The selected file is not a valid WebP image."};
    }
    var auth = "Basic " + Utilities.base64Encode(username + ":" + password);
    var endpoint = site + "/wp-json/wp/v2/media";
    var response = UrlFetchApp.fetch(endpoint, {
      method:"post",contentType:"image/webp",payload:bytes,
      headers:{"Authorization":auth,"Content-Disposition":'attachment; filename="'+filename+'"'},
      muteHttpExceptions:true
    });
    var status = response.getResponseCode();
    var media;
    try {media = JSON.parse(response.getContentText());} catch(parseError){media = {};}
    if (status < 200 || status >= 300 || !media.id || !media.source_url) {
      return {success:false,message:"WordPress upload failed (HTTP "+status+"): "+String(media.message || response.getContentText()).substring(0,350)};
    }
    uploadedUrl = String(media.source_url);
    // Store alt text in the WordPress media record.
    var metaResponse = UrlFetchApp.fetch(endpoint + "/" + encodeURIComponent(media.id),{
      method:"post",contentType:"application/json",
      payload:JSON.stringify({alt_text:altText}),
      headers:{"Authorization":auth},muteHttpExceptions:true
    });
    if (metaResponse.getResponseCode() < 200 || metaResponse.getResponseCode() >= 300) {
      return {success:false,url:uploadedUrl,mediaId:media.id,message:"Image uploaded to WordPress (media ID "+media.id+"), but alt text could not be set. URL shown below; check media details and save to sheet manually."};
    }
    var saved = saveFeaturedImageToSheet(uploadedUrl,filename,altText,postId);
    if (!saved.success) {
      return {success:false,url:uploadedUrl,mediaId:media.id,message:"Image uploaded to WordPress (media ID "+media.id+"), but sheet save failed: "+saved.message+" Use Save to Sheet to retry."};
    }
    return {success:true,url:uploadedUrl,mediaId:media.id,message:"Uploaded to WordPress (media ID "+media.id+"). "+saved.message};
  } catch(e) {
    return {success:false,url:uploadedUrl,message:"WordPress upload error: "+e.message+(uploadedUrl?" Image may have uploaded; check the media library.":"")};
  }
}
