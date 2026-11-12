<?php
/**
 * Abbey W6 Publishing Bridge — review/install as a standalone WordPress plugin.
 *
 * NOT DEPLOYED. Authenticated WordPress Application Password requests only.
 * Requires administrator and unfiltered_html capabilities.
 * Supports UPDATE only; REPUBLISH intentionally blocked pending syndication tests.
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
    if ($mode !== 'update') {
        return new WP_Error('w6_republish_disabled',
            'Republish is disabled until its syndication behaviour is verified.',
            array('status' => 400));
    }
    $post = get_post($post_id);
    if (!$post || $post->post_type !== 'post' || $post->post_status !== 'publish') {
        return new WP_Error('w6_post_not_found',
            'Expected an existing published WordPress post.',
            array('status' => 400));
    }
    $expected = (string) $request->get_param('expected_date');
    if (!$expected || $expected !== $post->post_date) {
        return new WP_Error('w6_date_conflict',
            'Post publication date changed since preflight. Reload the row.',
            array('status' => 409));
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
    if (!preg_match('~^<script\\s+type=["\\\']application/ld\\+json["\\\']\\s*>(.*?)</script>\\s*$~is', $schema, $m)) {
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

    // WordPress automatically advances post_modified; post_date is preserved.
    $updated_id = wp_update_post(wp_slash(array(
        'ID'           => $post_id,
        'post_title'   => $values['title'],
        'post_content' => $values['html'],
        'post_status'  => 'publish'
    )), true);
    if (is_wp_error($updated_id)) {
        return $updated_id;
    }

    update_post_meta($post_id, '_yoast_wpseo_title', wp_slash($values['meta_title']));
    update_post_meta($post_id, '_yoast_wpseo_metadesc', wp_slash($values['meta_description']));
    update_post_meta($post_id, '_yoast_wpseo_focuskw', wp_slash($values['focus_keyphrase']));
    update_post_meta($post_id, '_abbey_custom_schema', wp_slash($schema));
    clean_post_cache($post_id);

    $after = get_post($post_id);
    if (!$after || $after->post_date !== $expected) {
        return new WP_Error('w6_post_write_inconsistent',
            'WordPress saved the post but the publication date needs investigation.',
            array('status' => 500));
    }

    return rest_ensure_response(array(
        'success' => true,
        'post_id' => $post_id,
        'published_date' => $after->post_date,
        'modified_date' => $after->post_modified,
        'url' => get_permalink($post_id)
    ));
}
