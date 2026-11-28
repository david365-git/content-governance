<?php
/**
 * Plugin Name: Abbey Social Temporary Images
 * Description: Converts Abbey WordPress uploads into temporary JPEGs for reviewed social posts.
 * Version: 0.2.0
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
    if (!function_exists('imagecreatefromstring') || !function_exists('imagecreatetruecolor'))
        return new WP_Error('gd_required', 'PHP GD extension is required for contain/cover with white margins.', array('status'=>500));
    $maxw = max(400, min(1600, (int)($request->get_param('max_width') ?: 1200)));
    $maxh = max(400, min(1600, (int)($request->get_param('max_height') ?: 900)));
    $quality = max(75, min(95, (int)($request->get_param('quality') ?: 90)));
    $fit = strtoupper((string)($request->get_param('fit') ?: 'CONTAIN'));
    if (!in_array($fit, array('CONTAIN','COVER'), true))
        return new WP_Error('invalid_fit','Fit must be CONTAIN or COVER.',array('status'=>400));
    $background = strtoupper((string)($request->get_param('background') ?: '#FFFFFF'));
    if (!preg_match('/^#[0-9A-F]{6}$/', $background))
        return new WP_Error('invalid_background','Expected six-digit background colour.',array('status'=>400));
    $fingerprint = hash('sha256', $source.'|'.filemtime($source).'|'.$maxw.'x'.$maxh.'|'.$quality.'|'.$fit.'|'.$background.'|v2');
    $filename = 'social-'.$fingerprint.'.jpg';
    $dest = trailingslashit($paths['dir']).$filename;
    $meta = $dest.'.json';
    if (!is_file($dest)) {
        $details = @getimagesize($source);
        if (!$details || $details[0] < 1 || $details[1] < 1)
            return new WP_Error('invalid_image','Source image dimensions unavailable.',array('status'=>400));
        if ($details[0] * $details[1] > 30000000 || @filesize($source) > 30000000)
            return new WP_Error('source_too_large','Image too large for safe resizing.',array('status'=>400));
        $bytes = @file_get_contents($source);
        if ($bytes === false) return new WP_Error('source_unreadable','Image could not be read.',array('status'=>400));
        $src = @imagecreatefromstring($bytes);
        unset($bytes);
        if (!$src) return new WP_Error('unsupported_image','PHP GD could not decode the image.',array('status'=>400));
        $sw = imagesx($src); $sh = imagesy($src);
        $canvas = imagecreatetruecolor($maxw,$maxh);
        if (!$canvas) { imagedestroy($src); return new WP_Error('canvas_failed','Could not prepare canvas.',array('status'=>500)); }
        $r = hexdec(substr($background,1,2));
        $g = hexdec(substr($background,3,2));
        $b = hexdec(substr($background,5,2));
        $white = imagecolorallocate($canvas,$r,$g,$b);
        imagefilledrectangle($canvas,0,0,$maxw,$maxh,$white);
        $scale = $fit === 'COVER' ? max($maxw/$sw,$maxh/$sh) : min(1,$maxw/$sw,$maxh/$sh);
        $dw = max(1,(int)round($sw*$scale)); $dh = max(1,(int)round($sh*$scale));
        $x = (int)floor(($maxw-$dw)/2); $y = (int)floor(($maxh-$dh)/2);
        $copied = imagecopyresampled($canvas,$src,$x,$y,0,0,$dw,$dh,$sw,$sh);
        imagedestroy($src);
        if (!$copied || !imagejpeg($canvas,$dest,$quality)) {
            imagedestroy($canvas);
            return new WP_Error('conversion_failed','Could not save platform-specific JPEG.',array('status'=>500));
        }
        imagedestroy($canvas);
        @chmod($dest,0644);
    }
    file_put_contents($meta,wp_json_encode(array('released_at'=>null,'prepared_at'=>time())));
    return rest_ensure_response(array(
        'jpg_url'=>trailingslashit($paths['url']).$filename,
        'token'=>$fingerprint,'width'=>$maxw,'height'=>$maxh,'status'=>'READY',
        'fit'=>$fit,'background'=>$background
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
