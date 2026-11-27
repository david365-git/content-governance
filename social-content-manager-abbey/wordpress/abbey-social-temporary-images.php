<?php
/**
 * Plugin Name: Abbey Social Temporary Images
 * Description: Converts Abbey WordPress uploads into temporary JPEGs for reviewed social posts.
 * Version: 0.1.0
 */
if (!defined('ABSPATH')) exit;

add_action('rest_api_init', function () {
    register_rest_route('abbey-social/v1', '/prepare', array(
        'methods' => 'POST',
        'callback' => 'abbey_social_prepare_image',
        'permission_callback' => function () { return current_user_can('upload_files'); },
    ));
    register_rest_route('abbey-social/v1', '/release', array(
        'methods' => 'POST',
        'callback' => 'abbey_social_release_image',
        'permission_callback' => function () { return current_user_can('upload_files'); },
    ));
});
add_action('init', function () {
    if (!wp_next_scheduled('abbey_social_cleanup')) {
        wp_schedule_event(time() + HOUR_IN_SECONDS, 'daily', 'abbey_social_cleanup');
    }
});
register_deactivation_hook(__FILE__, function () {
    wp_clear_scheduled_hook('abbey_social_cleanup');
});
function abbey_social_paths() {
    $up = wp_get_upload_dir();
    if (!empty($up['error'])) return new WP_Error('uploads_unavailable', 'Uploads directory unavailable', array('status'=>500));
    $dir = trailingslashit($up['basedir']) . 'abbey-social-temp';
    if (!wp_mkdir_p($dir)) return new WP_Error('temp_unavailable', 'Temporary image directory unavailable', array('status'=>500));
    return array('basedir'=>$up['basedir'], 'baseurl'=>$up['baseurl'], 'dir'=>$dir,
        'url'=>trailingslashit($up['baseurl']).'abbey-social-temp');
}
function abbey_social_source($url, $paths) {
    $url = esc_url_raw($url);
    $base = trailingslashit($paths['baseurl']);
    if (strpos($url, $base) !== 0) return new WP_Error('invalid_source', 'Only images in this WordPress uploads directory are allowed', array('status'=>400));
    $relative = rawurldecode(substr($url, strlen($base)));
    if (strpos($relative, '?') !== false) $relative = explode('?', $relative, 2)[0];
    if (strpos($relative, '#') !== false) $relative = explode('#', $relative, 2)[0];
    if (strpos($relative, 'abbey-social-temp/') === 0) return new WP_Error('invalid_source', 'Temporary images cannot be source images', array('status'=>400));
    $real = realpath(trailingslashit($paths['basedir']) . $relative);
    $uploads_real = realpath($paths['basedir']);
    if (!$real || !$uploads_real || strpos($real, trailingslashit($uploads_real)) !== 0 || !is_file($real))
        return new WP_Error('invalid_source', 'Source image does not exist inside uploads', array('status'=>400));
    $type = wp_check_filetype($real);
    if (!in_array(strtolower($type['ext']), array('webp','jpg','jpeg','png'), true))
        return new WP_Error('unsupported_source', 'Only WebP, JPEG or PNG source images are accepted', array('status'=>400));
    return $real;
}
function abbey_social_prepare_image(WP_REST_Request $request) {
    $paths = abbey_social_paths();
    if (is_wp_error($paths)) return $paths;
    $source = abbey_social_source((string)$request->get_param('url'), $paths);
    if (is_wp_error($source)) return $source;
    $maxw = max(400, min(1600, (int)($request->get_param('max_width') ?: 1200)));
    $maxh = max(400, min(1600, (int)($request->get_param('max_height') ?: 1200)));
    $quality = max(75, min(95, (int)($request->get_param('quality') ?: 90)));
    $fingerprint = hash('sha256', $source.'|'.filemtime($source).'|'.$maxw.'x'.$maxh.'|'.$quality);
    $filename = 'social-'.$fingerprint.'.jpg';
    $dest = trailingslashit($paths['dir']).$filename;
    $meta = $dest.'.json';
    if (!is_file($dest)) {
        $editor = wp_get_image_editor($source);
        if (is_wp_error($editor)) return new WP_Error('conversion_failed', $editor->get_error_message(), array('status'=>500));
        $editor->resize($maxw, $maxh, false);
        $editor->set_quality($quality);
        $result = $editor->save($dest, 'image/jpeg');
        if (is_wp_error($result)) return new WP_Error('conversion_failed', $result->get_error_message(), array('status'=>500));
        @chmod($dest, 0644);
    }
    // A prepared file remains available until it is explicitly released.
    file_put_contents($meta, wp_json_encode(array('released_at'=>null,'prepared_at'=>time())));
    return rest_ensure_response(array(
        'jpg_url'=>trailingslashit($paths['url']).$filename,
        'token'=>$fingerprint, 'width'=>@getimagesize($dest)[0],
        'height'=>@getimagesize($dest)[1], 'status'=>'READY'
    ));
}
function abbey_social_release_image(WP_REST_Request $request) {
    $paths = abbey_social_paths();
    if (is_wp_error($paths)) return $paths;
    $token = (string)$request->get_param('token');
    if (!preg_match('/^[a-f0-9]{64}$/', $token)) return new WP_Error('invalid_token','Invalid image token',array('status'=>400));
    $file = trailingslashit($paths['dir']).'social-'.$token.'.jpg';
    if (!is_file($file)) return new WP_Error('not_found','Temporary JPG not found',array('status'=>404));
    $meta = $file.'.json';
    file_put_contents($meta, wp_json_encode(array('released_at'=>time())));
    return rest_ensure_response(array('status'=>'RELEASED','deletes_after_days'=>7));
}
add_action('abbey_social_cleanup', function () {
    $paths = abbey_social_paths();
    if (is_wp_error($paths)) return;
    foreach (glob(trailingslashit($paths['dir']).'social-*.jpg.json') ?: array() as $meta) {
        $info = json_decode((string)@file_get_contents($meta), true);
        if (!is_array($info) || empty($info['released_at'])) continue;
        if ((int)$info['released_at'] > time() - 7 * DAY_IN_SECONDS) continue;
        $jpg = substr($meta, 0, -5);
        if (is_file($jpg)) unlink($jpg);
        unlink($meta);
    }
});
