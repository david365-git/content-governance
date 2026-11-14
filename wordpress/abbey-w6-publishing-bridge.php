<?php
/**
 * Plugin Name: Abbey W6 Publishing Bridge
 * Description: Authenticated update-only publishing endpoint for the Abbey Content Engine.
 * Version: 0.2.0
 * Requires PHP: 7.4
 * Abbey W6 Publishing Bridge — review/install as a standalone WordPress plugin.
 *
 * NOT DEPLOYED. Authenticated WordPress Application Password requests only.
 * Requires administrator and unfiltered_html capabilities.
 * UPDATE preserves publication date; REPUBLISH resets it with a 30-day safeguard.
 * Custom schema stays in _abbey_custom_schema, consumed by existing output rules.
 */
defined('ABSPATH') || exit;

add_action('rest_api_init', function () {
    register_rest_route('abbey/v1', '/w6-publish', array(
        'methods' => WP_REST_Server::CREATABLE,
        'callback' => 'abbey_w6_publish_post',
        'permission_callback' => function (WP_REST_Request $request) {
            $post_id = absint($request->get_param('post_id'));
            return $post_id > 0
                && current_user_can('manage_options')
                && current_user_can('edit_post', $post_id)
                && current_user_can('unfiltered_html');
        }
    ));
});

function abbey_w6_publish_post(WP_REST_Request $request) {
    $post_id = absint($request->get_param('post_id'));
    $mode = (string) $request->get_param('mode');
    if (!in_array($mode, array('update', 'republish'), true)) {
        return new WP_Error('w6_invalid_mode',
            'Choose Update or Republish.', array('status' => 400));
    }
    $post = get_post($post_id);
    if (!$post || $post->post_type !== 'post' || $post->post_status !== 'publish') {
        return new WP_Error('w6_post_not_found',
            'Expected an existing published WordPress post.',
            array('status' => 400));
    }
    // REST API date uses ISO "T"; WP_Post::post_date uses a space.
    // Compare only equivalent site-local timestamps, never change post_date.
    $expected = (string) $request->get_param('expected_date');
    if (!preg_match('/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/D', $expected) ||
        str_replace('T', ' ', $expected) !== $post->post_date) {
        return new WP_Error('w6_date_conflict',
            'Post publication date changed since preflight. Reload the row.',
            array('status' => 409));
    }

    if ($mode === 'republish') {
        // A second safeguard: do not republish recent or future-dated articles.
        // This check is independent of the spreadsheet's grey highlighting.
        $published_at = DateTimeImmutable::createFromFormat(
            '!Y-m-d H:i:s', $post->post_date, wp_timezone());
        $now_at = new DateTimeImmutable('now', wp_timezone());
        if (!$published_at || $published_at > $now_at ||
            ($now_at->getTimestamp() - $published_at->getTimestamp()) < 30 * DAY_IN_SECONDS) {
            return new WP_Error('w6_recent_or_future_publication',
                'Republish blocked: the publication date is within the last 30 days or scheduled in the future.',
                array('status' => 409));
        }
        $expected_modified = (string) $request->get_param('expected_modified');
        if (!$expected_modified ||
            str_replace('T', ' ', $expected_modified) !== $post->post_modified) {
            return new WP_Error('w6_modified_conflict',
                'WordPress changed since verification. Verify the post again before republishing.',
                array('status' => 409));
        }
    }

    $fields = array('title','html','meta_title','meta_description','focus_keyphrase','schema');
    $values = array();
    foreach ($fields as $key) {
        $value = $request->get_param($key);
        if (!is_string($value)) {
            return new WP_Error('w6_missing_field',
                'Missing text field: ' . $key, array('status' => 400));
        }
        $values[$key] = $value;
    }
    if (trim($values['title']) === '' || trim($values['html']) === '') {
        return new WP_Error('w6_empty_content',
            'Title and HTML must not be empty.', array('status' => 400));
    }
    if (strlen($values['html']) > 1500000) {
        return new WP_Error('w6_content_too_large',
            'HTML exceeds the allowed size.', array('status' => 400));
    }

    // Existing importer writes one <script type="application/ld+json"> payload.
    // Validate that structure and reject any executable non-JSON script.
    $schema = trim($values['schema']);
    if (!preg_match('~^<script\\s+type="application/ld\\+json"\\s*>\\s*(.*?)\\s*</script>\\s*$~is', $schema, $m)) {
        return new WP_Error('w6_invalid_schema',
            'Schema must be one JSON-LD script block.', array('status' => 400));
    }
    $decoded = json_decode($m[1], true);
    if (json_last_error() !== JSON_ERROR_NONE || !is_array($decoded)
        || !isset($decoded['@context'])
        || $decoded['@context'] !== 'https://schema.org'
        || (!isset($decoded['@graph']) && !isset($decoded['@type']))) {
        return new WP_Error('w6_invalid_schema_json',
            'Schema must contain valid schema.org JSON-LD.',
            array('status' => 400));
    }

    // Optional featured image URL supplied from posts column EM.
    // Resolve it to an existing Media Library image BEFORE changing article content.
    // Blank or omitted means leave featured_media exactly as it is.
    $feature_action = 'unchanged';
    $feature_id = null;
    if ($request->has_param('featured_image_url')) {
        $image_url = trim((string) $request->get_param('featured_image_url'));
        if ($image_url !== '') {
            $parsed_url = wp_parse_url($image_url);
            $home_host = wp_parse_url(home_url(), PHP_URL_HOST);
            if (!is_array($parsed_url) ||
                strtolower((string) ($parsed_url['scheme'] ?? '')) !== 'https' ||
                strtolower((string) ($parsed_url['host'] ?? '')) !== strtolower((string) $home_host)) {
                return new WP_Error('w6_image_domain',
                    'Column EM must contain an HTTPS image URL on this WordPress website.',
                    array('status' => 400));
            }
            $feature_id = absint(attachment_url_to_postid($image_url));
            if (!$feature_id || get_post_type($feature_id) !== 'attachment' ||
                !wp_attachment_is_image($feature_id)) {
                return new WP_Error('w6_image_not_found',
                    'Column EM image URL was not found as an image in the WordPress Media Library. No article changes made.',
                    array('status' => 400));
            }
            if ((int) get_post_thumbnail_id($post_id) !== $feature_id) {
                $feature_action = 'updated';
            }
        }
    }

    // Update leaves the publication date untouched. Republish explicitly
    // advances the publication date through the normal WordPress post API.
    $post_fields = array(
        'ID'           => $post_id,
        'post_title'   => $values['title'],
        'post_content' => $values['html'],
        'post_status'  => 'publish'
    );
    if ($mode === 'republish') {
        $post_fields['post_date'] = current_time('mysql');
        $post_fields['post_date_gmt'] = get_gmt_from_date($post_fields['post_date']);
    }
    $updated_id = wp_update_post(wp_slash($post_fields), true);
    if (is_wp_error($updated_id)) {
        return $updated_id;
    }

    update_post_meta($post_id, '_yoast_wpseo_title', wp_slash($values['meta_title']));
    update_post_meta($post_id, '_yoast_wpseo_metadesc', wp_slash($values['meta_description']));
    update_post_meta($post_id, '_yoast_wpseo_focuskw', wp_slash($values['focus_keyphrase']));
    update_post_meta($post_id, '_abbey_custom_schema', wp_slash($schema));
    if ($feature_id !== null && $feature_action === 'updated') {
        if (!set_post_thumbnail($post_id, $feature_id)) {
            return new WP_Error('w6_feature_image_failed',
                'Article saved but featured image could not be set. Check the post before retrying.',
                array('status' => 500));
        }
    }
    clean_post_cache($post_id);

    $after = get_post($post_id);
    if (!$after ||
        ($mode === 'update' && $after->post_date !== str_replace('T', ' ', $expected)) ||
        ($mode === 'republish' && $after->post_date === str_replace('T', ' ', $expected))) {
        return new WP_Error('w6_post_write_inconsistent',
            'WordPress saved the post but the publication date needs investigation.',
            array('status' => 500));
    }

    return rest_ensure_response(array(
        'success' => true,
        'post_id' => $post_id,
        'mode' => $mode,
        'published_date' => str_replace(' ', 'T', $after->post_date),
        'modified_date' => str_replace(' ', 'T', $after->post_modified),
        'url' => get_permalink($post_id),
        'featured_image_action' => $feature_action,
        'featured_image_url' => get_the_post_thumbnail_url($post_id, 'full') ?: ''
    ));
}
