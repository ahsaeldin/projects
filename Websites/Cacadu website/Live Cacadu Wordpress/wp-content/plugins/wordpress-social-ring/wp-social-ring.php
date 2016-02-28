<?php
/*
Plugin Name: WordPress Social Ring (Facebook Like, Google +1, ReTweet and Pin It)
Description: Let visitors share posts/pages on Facebook, Twitter and Google+. From admin page you can choose which button display: Facebook Like, Facebook Send, Facebook Share, Google +1 and Twitter.
Author: Niccol&ograve; Tapparo
Version: 1.1.8
Author URI: http://wordpress.altervista.org/
Plugin URI: http://wordpress.altervista.org/en/tag/wordpress-social-ring-en/
*/

define( 'WP_SOCIAL_RING', 'wp_social_ring' );
define( 'WP_SOCIAL_RING_PATH', plugin_dir_path(__FILE__) );
define( 'WP_SOCIAL_RING_URL', plugin_dir_url(__FILE__) );

register_activation_hook(__FILE__,'social_ring_install');

load_plugin_textdomain(WP_SOCIAL_RING, false, dirname(plugin_basename(__FILE__)).'/langs/');

//Set defaults if not defined
$wp_social_ring_options = get_option(WP_SOCIAL_RING.'_options');

include WP_SOCIAL_RING_PATH.'/includes/library.php';
include WP_SOCIAL_RING_PATH.'/includes/widgets.php';

if(is_admin()) {
	include WP_SOCIAL_RING_PATH.'/admin/admin.php';
}

add_filter('language_attributes','social_ring_add_opengraph_namespace');
add_action('wp_head', 'social_ring_add_opengraph_meta');
add_action('wp_head', 'social_ring_add_css');
add_filter('the_content', 'social_ring_add_sharing');
add_action('wp_footer', 'social_ring_add_js');
add_shortcode('socialring', 'social_ring_shortcode');

function social_ring_install() {

	if(version_compare(get_bloginfo('version'), '3.2', '<')) {
		deactivate_plugins(basename(__FILE__));
	} else {
		$wp_social_ring_options = array (
			'social_facebook_like_button' => 0,
			'social_facebook_send_button' => 0,
			'social_facebook_share_button' => 0,
			'social_twitter_button' => 0,
			'social_google_button' => 0,
			'social_pin_it_button' => 0,
			'social_on_home' => 0,
			'social_on_pages' => 0,
			'social_on_posts' => 0,
			'social_on_category' => 0,
			'social_on_archive' => 0,
			'social_before_content' => 0,
			'social_after_content' => 0,
			'opengraph' => 0,
			'social_cross_plugin_compatibility' => 0,
		);
		update_option(WP_SOCIAL_RING.'_options', $wp_social_ring_options);
	}
}


?>