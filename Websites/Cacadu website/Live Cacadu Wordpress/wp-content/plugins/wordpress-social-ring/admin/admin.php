<?php

function wp_social_ring_setting_page() {
	
	?>
	
	<div class="wrap">
		<?php screen_icon('plugins'); ?>
		<h2><?php _e('WordPress Social Ring Settings',WP_SOCIAL_RING); ?></h2>
		<div id="wp-social-ring">
			<div class="postbox-container" style="width:70%;">
				<form action="options.php" method="post">
				
						<div class="postbox">
							<?php settings_fields(WP_SOCIAL_RING.'_options'); ?>
							<?php do_settings_sections(WP_SOCIAL_RING); ?>
						</div>
						<div class="postbox">
							<h3><?php _e('Extra',WP_SOCIAL_RING); ?></h3>
							<div style="margin-left:10px;">
								<p style="margin-top:10px;">
									<?php _e('<b>Custom position</b>',WP_SOCIAL_RING); ?>
								</p>	
								<p>
									<?php _e('You can give Social buttons a custom position by calling this function in your theme.',WP_SOCIAL_RING); ?>
								</p>
								<pre>&lt;?php if(function_exists('social_ring_show')){ social_ring_show();} ?&gt;</pre>
								<p>
									<?php _e('Pay Attention: it must called <b>inside the loop</b> to work properly!',WP_SOCIAL_RING); ?>
								</p>
								<p style="margin-top:20px;">
									<?php _e('<b>Shortcode</b>',WP_SOCIAL_RING); ?>
								</p>
								<p>
									<?php _e('You can also place social buttons inside your posts by using the shortcode:',WP_SOCIAL_RING); ?>
								</p>
								<pre style="margin-bottom:20px;">[socialring]</pre>
							</div>
						</div>
						<input name="submit" class="button-primary" type="submit" value="<?php _e('Save Changes',WP_SOCIAL_RING); ?>" />
				
				</form>
			</div>
			<div class="postbox-container" style="margin-left:15px;">
				<div>
					<a target="_blank" href="<?php _e('http://en.altervista.org/create-free-blog.html?ref=socialring',WP_SOCIAL_RING); ?>">
								<img alt="<?php _e('Create your free blog!');?>" title="<?php _e('Create your free blog!');?>" src="<?php echo WP_SOCIAL_RING_URL; ?>/admin/images/altervista.gif" height="300" width="180" />
					</a>
				</div>
				<div class="postbox" style="min-width:180px !important;margin-top:5px;text-align:center;padding: 8px 0;">
					<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
						<input type="hidden" name="cmd" value="_s-xclick">
						<input type="hidden" name="encrypted" value="-----BEGIN PKCS7-----MIIHRwYJKoZIhvcNAQcEoIIHODCCBzQCAQExggEwMIIBLAIBADCBlDCBjjELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAkNBMRYwFAYDVQQHEw1Nb3VudGFpbiBWaWV3MRQwEgYDVQQKEwtQYXlQYWwgSW5jLjETMBEGA1UECxQKbGl2ZV9jZXJ0czERMA8GA1UEAxQIbGl2ZV9hcGkxHDAaBgkqhkiG9w0BCQEWDXJlQHBheXBhbC5jb20CAQAwDQYJKoZIhvcNAQEBBQAEgYCKUi69ZExrjK6f0vt1msE9HX8NESqdbNsoFtBAp0knIRdACrynyki9tcuFfxJFTkdswBEyE9jwxDo2UOIAskQGlAo7bUvz8VwlrO6qwjhBYQtHYX5we9tdKI9WA08Sj2QA63XKMpZ2xbSXVb2nGMoCvIFa75PpFWgf+llFbW8P7jELMAkGBSsOAwIaBQAwgcQGCSqGSIb3DQEHATAUBggqhkiG9w0DBwQIN7Wyr2UpdmCAgaAnNkTfpNYNC9bHOWJ+xt/+hQ0lSW+J4WjUIbN57xIIgsdR6dBFNIjeZ32G2HN2djy5tJxXKfuoJm+KkKw6aCD0o0BFZCwMUcJqvvX5YCXYJ4jXb5KzlhV8h6KlkEdu51tCiEbiKQotRoEFrIlhFro0w9SNhMiKMS6rNrqW6amFdhGUHKSQdhaRcZ43NzPtHWvR9qjy3B+MWIRHTz7+bjOtoIIDhzCCA4MwggLsoAMCAQICAQAwDQYJKoZIhvcNAQEFBQAwgY4xCzAJBgNVBAYTAlVTMQswCQYDVQQIEwJDQTEWMBQGA1UEBxMNTW91bnRhaW4gVmlldzEUMBIGA1UEChMLUGF5UGFsIEluYy4xEzARBgNVBAsUCmxpdmVfY2VydHMxETAPBgNVBAMUCGxpdmVfYXBpMRwwGgYJKoZIhvcNAQkBFg1yZUBwYXlwYWwuY29tMB4XDTA0MDIxMzEwMTMxNVoXDTM1MDIxMzEwMTMxNVowgY4xCzAJBgNVBAYTAlVTMQswCQYDVQQIEwJDQTEWMBQGA1UEBxMNTW91bnRhaW4gVmlldzEUMBIGA1UEChMLUGF5UGFsIEluYy4xEzARBgNVBAsUCmxpdmVfY2VydHMxETAPBgNVBAMUCGxpdmVfYXBpMRwwGgYJKoZIhvcNAQkBFg1yZUBwYXlwYWwuY29tMIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDBR07d/ETMS1ycjtkpkvjXZe9k+6CieLuLsPumsJ7QC1odNz3sJiCbs2wC0nLE0uLGaEtXynIgRqIddYCHx88pb5HTXv4SZeuv0Rqq4+axW9PLAAATU8w04qqjaSXgbGLP3NmohqM6bV9kZZwZLR/klDaQGo1u9uDb9lr4Yn+rBQIDAQABo4HuMIHrMB0GA1UdDgQWBBSWn3y7xm8XvVk/UtcKG+wQ1mSUazCBuwYDVR0jBIGzMIGwgBSWn3y7xm8XvVk/UtcKG+wQ1mSUa6GBlKSBkTCBjjELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAkNBMRYwFAYDVQQHEw1Nb3VudGFpbiBWaWV3MRQwEgYDVQQKEwtQYXlQYWwgSW5jLjETMBEGA1UECxQKbGl2ZV9jZXJ0czERMA8GA1UEAxQIbGl2ZV9hcGkxHDAaBgkqhkiG9w0BCQEWDXJlQHBheXBhbC5jb22CAQAwDAYDVR0TBAUwAwEB/zANBgkqhkiG9w0BAQUFAAOBgQCBXzpWmoBa5e9fo6ujionW1hUhPkOBakTr3YCDjbYfvJEiv/2P+IobhOGJr85+XHhN0v4gUkEDI8r2/rNk1m0GA8HKddvTjyGw/XqXa+LSTlDYkqI8OwR8GEYj4efEtcRpRYBxV8KxAW93YDWzFGvruKnnLbDAF6VR5w/cCMn5hzGCAZowggGWAgEBMIGUMIGOMQswCQYDVQQGEwJVUzELMAkGA1UECBMCQ0ExFjAUBgNVBAcTDU1vdW50YWluIFZpZXcxFDASBgNVBAoTC1BheVBhbCBJbmMuMRMwEQYDVQQLFApsaXZlX2NlcnRzMREwDwYDVQQDFAhsaXZlX2FwaTEcMBoGCSqGSIb3DQEJARYNcmVAcGF5cGFsLmNvbQIBADAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTExMTI5MTMxNDMwWjAjBgkqhkiG9w0BCQQxFgQUjUOP1N4V9c0h6YAMMzgjddXAO4MwDQYJKoZIhvcNAQEBBQAEgYBZw6WfFp43NACMckt0OC4yRhqpo03hml14U3Y4eP6XExKr1pNIlhcGvKloDXWmdiydEtpza5R+ZE3D4e42B/BupJ5sL2Q4xCIZmmOm9as0Rt44+xRXXIXIjJq6PYdxbfReDSRFfEC6jO+zfVu12SlBQgkiwZqWXWmc1VBFFQYwIA==-----END PKCS7-----
						">
						<input type="image" src="https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif" border="0" name="submit" alt="PayPal - The safer, easier way to pay online!">
						<img alt="" border="0" src="https://www.paypalobjects.com/it_IT/i/scr/pixel.gif" width="1" height="1">
					</form>
				</div>
			</div>
		</div>
	</div>
	<?php

}


function add_wp_social_ring_css_js() {
	
	wp_enqueue_style( WP_SOCIAL_RING.'-style', WP_SOCIAL_RING_URL.'admin/css/style.css');
	wp_enqueue_script( WP_SOCIAL_RING.'_facebook_js', 'http://connect.facebook.net/'.__('en_US', WP_SOCIAL_RING).'/all.js#xfbml=1', array(), false, true);
	wp_enqueue_script( WP_SOCIAL_RING.'_gplus_js', 'https://apis.google.com/js/plusone.js', array(), false, true);
	wp_enqueue_script( WP_SOCIAL_RING.'_twitter_js', 'http://platform.twitter.com/widgets.js', array(), false, true);
	wp_enqueue_script( WP_SOCIAL_RING.'_pin_it_js', 'http://assets.pinterest.com/js/pinit.js', array(), false, true);
}

// Register and define the settings
add_action('admin_init', 'wp_social_ring_admin_init');
function wp_social_ring_admin_init(){

	

	register_setting(
		WP_SOCIAL_RING.'_options',
		WP_SOCIAL_RING.'_options',
		WP_SOCIAL_RING.'_validate_options'
	);

	add_settings_section(
		WP_SOCIAL_RING.'_setting_section',
		__('General Settings',WP_SOCIAL_RING),
		'wp_social_ring_social_share_explain',
		WP_SOCIAL_RING
	);
	add_settings_field(
		'wp_social_ring_social_buttons',
		__('Buttons',WP_SOCIAL_RING),
		'print_social_ring_buttons_input',
		WP_SOCIAL_RING,
		WP_SOCIAL_RING.'_setting_section'
	);
	add_settings_field(
		'wp_social_ring_position',
		__('Position',WP_SOCIAL_RING),
		'print_social_ring_position_input',
		WP_SOCIAL_RING,
		WP_SOCIAL_RING.'_setting_section'
	);
	add_settings_field(
		'wp_social_ring_show_on',
		__('Show on',WP_SOCIAL_RING),
		'print_social_ring_show_on_input',
		WP_SOCIAL_RING,
		WP_SOCIAL_RING.'_setting_section'
	);
}

function wp_social_ring_social_share_explain() {
	?>
		<div class="explain"><?php _e('Choose Social Ring position and behavior',WP_SOCIAL_RING); ?></div>
	<?php
}


function print_social_ring_position_input() {
	
	global $wp_social_ring_options;
	// echo the field
	?>
		
		<ul>
			<li>
				<span><input id='social_before_content' name='wp_social_ring_options[social_before_content]' type='checkbox' value="1" <?php if($wp_social_ring_options['social_before_content'] == 1) echo "checked"; ?> /></span>
				<span><?php _e('Before content',WP_SOCIAL_RING) ?></span>
			</li>
			<li>
				<span><input id='social_after_content' name='wp_social_ring_options[social_after_content]' type='checkbox' value="1" <?php if($wp_social_ring_options['social_after_content'] == 1) echo "checked"; ?> /></span>
				<span><?php _e('After content',WP_SOCIAL_RING) ?></span>
			</li>
		</ul>

	<?php
}

function print_social_ring_buttons_input() {
	
	global $wp_social_ring_options;
	// echo the field
	?>
		<ul>
			<li style="clear:both;">
				<div style="float:left;"><input id='social_facebook_like_button' name='wp_social_ring_options[social_facebook_like_button]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_facebook_like_button'], 1); ?> /></div>
				<div style="float:left;margin-left:20px;width:150px;"><?php _e('Facebook Like',WP_SOCIAL_RING) ?></div>
				<div style="float:left;margin-left:20px;"><fb:like href="http://www.facebook.com/pages/DrWordPress/166397626712895" send="false" showfaces="false" width="108" layout="button_count" action="like"/></fb:like></div>
			</li>
			<li style="clear:both;padding-top:10px;">
				<div style="float:left;"><input id='social_facebook_send_button' name='wp_social_ring_options[social_facebook_send_button]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_facebook_send_button'], 1); ?> /></div>
				<div style="float:left;margin-left:20px;width:150px;"><?php _e('Facebook Send',WP_SOCIAL_RING) ?></div>
				<div style="float:left;margin-left:20px;"><fb:like href="http://www.facebook.com/pages/DrWordPress/166397626712895" send="true" showfaces="false" width="164" layout="button_count" action="like"/></fb:like></div>
				<div id="fb-root"></div>
			</li>
			<li style="clear:both;padding-top:10px;">
				<div style="float:left;"><input id='social_facebook_share_button' name='wp_social_ring_options[social_facebook_share_button]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_facebook_share_button'], 1); ?> /></div>
				<div style="float:left;margin-left:20px;width:150px;"><?php _e('Facebook Share',WP_SOCIAL_RING) ?></div>
				<div style="float:left;margin-left:20px;"><iframe allowtransparency="true" frameborder="0" hspace="0" marginheight="0" marginwidth="0" scrolling="no" style="width: 58px; height: 21px; position: static; left: 0px; top: 0px; visibility: visible; " tabindex="-1" vspace="0" width="100%" src="<?php echo WP_SOCIAL_RING_URL; ?>includes/share.php?url=<?php echo urlencode('http://wordpress.altervista.org/'); ?>"></iframe></div>
			</li>
			<li style="clear:both;padding-top:10px;">
				<div style="float:left;"><input id='social_twitter_button' name='wp_social_ring_options[social_twitter_button]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_twitter_button'], 1); ?> /></div>
				<div style="float:left;margin-left:20px;width:150px;"><?php _e('Twitter',WP_SOCIAL_RING) ?></div>
				<div style="float:left;margin-left:20px;"><a href="http://twitter.com/share" data-url="<?php echo site_url(); ?>" data-text="<?php _e('I use WordPress Social Ring on my blog', WP_SOCIAL_RING); ?>" data-via="dottorwordpress" data-count="horizontal" class="sr-twitter-button twitter-share-button"></a></div>
			</li>
			<li style="clear:both;padding-top:10px;">
				<div style="float:left;"><input id='social_google_button' name='wp_social_ring_options[social_google_button]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_google_button'], 1); ?> /></div>
				<div style="float:left;margin-left:20px;width:150px;"><?php _e('Google +1',WP_SOCIAL_RING) ?></div>
				<div style="float:left;margin-left:20px;"><g:plusone href="<?php _e('http://wordpress.altervista.org/', WP_SOCIAL_RING); ?>" size="medium" callback="plusone_vote"></g:plusone></div>
			</li>
			
			<li style="clear:both;padding-top:10px;">
				<div style="float:left;"><input id='social_linkedin_button' name='wp_social_ring_options[social_linkedin_button]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_linkedin_button'], 1); ?> /></div>
				<div style="float:left;margin-left:20px;width:150px;"><?php _e('LinkedIn',WP_SOCIAL_RING) ?></div>
				<div style="float:left;margin-left:20px;"><div class="social-ring-button"><script src="//platform.linkedin.com/in.js" type="text/javascript"></script><script type="IN/Share" data-url="http://wordpress.altervista.org/" data-counter="right"></script></div>
			</li>
			
			<li style="clear:both;padding-top:10px;">
				<div style="float:left;"><input id='social_pin_it_button' name='wp_social_ring_options[social_pin_it_button]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_pin_it_button'], 1); ?> /></div>
				<div style="float:left;margin-left:20px;width:150px;"><?php _e('Pin it',WP_SOCIAL_RING) ?></div>
				<div style="float:left;margin-left:20px;"><a href="http://pinterest.com/pin/create/button/?url=<?php echo urlencode(__('http://wordpress.altervista.org/', WP_SOCIAL_RING)); ?>&media=<?php echo urlencode('http://wordpress.altervista.org/logo.jpg'); ?>&description=<?php echo urlencode(__('WordPress tutorials, plugin and themes.', WP_SOCIAL_RING)." ".site_url()); ?>" class="pin-it-button" count-layout="horizontal"></a></div>
			</li>
		</ul>
	<?php
}

function print_social_ring_show_on_input() {
	
	global $wp_social_ring_options;
	// echo the field

	?>

		<ul>
			<li>
				<span><input id='social_on_home' name='wp_social_ring_options[social_on_home]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_on_home'], 1); ?> /></span>
				<span><?php _e('Home',WP_SOCIAL_RING) ?></span>
			</li>
			<li>
				<span><input id='social_on_pages' name='wp_social_ring_options[social_on_pages]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_on_pages'], 1); ?> /></span>
				<span><?php _e('Pages',WP_SOCIAL_RING) ?></span>
			</li>
			<li>
				<span><input id='social_on_posts' name='wp_social_ring_options[social_on_posts]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_on_posts'], 1); ?> /></span>
				<span><?php _e('Posts',WP_SOCIAL_RING) ?></span>
			</li>
			<li>
				<span><input id='social_on_category' name='wp_social_ring_options[social_on_category]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_on_category'], 1); ?> /></span>
				<span><?php _e('Categories',WP_SOCIAL_RING) ?></span>
			</li>
			<li>
				<span><input id='social_on_archive' name='wp_social_ring_options[social_on_archive]' type='checkbox' value="1" <?php checked($wp_social_ring_options['social_on_archive'], 1); ?> /></span>
				<span><?php _e('Archive',WP_SOCIAL_RING) ?></span>
			</li>
		</ul>

	<?php
}


// Validate user input (we want text only)
function wp_social_ring_validate_options( $input ) {
	
	global $wp_social_ring_options;
	
	//social
	$valid['social_facebook_like_button'] = (isset( $input['social_facebook_like_button'])) ? 1 : 0;
	$valid['social_facebook_send_button'] = (isset( $input['social_facebook_send_button'])) ? 1 : 0;
	if($valid['social_facebook_send_button'] == 1) {
		$valid['social_facebook_like_button'] = 1;
	}
	$valid['social_facebook_share_button'] = (isset( $input['social_facebook_share_button'])) ? 1 : 0;
	$valid['social_twitter_button'] = (isset( $input['social_twitter_button'])) ? 1 : 0;
	$valid['social_google_button'] = (isset( $input['social_google_button'])) ? 1 : 0;
	$valid['social_pin_it_button'] = (isset( $input['social_pin_it_button'])) ? 1 : 0;
	$valid['social_linkedin_button'] = (isset( $input['social_linkedin_button'])) ? 1 : 0;
	$valid['social_on_home'] = (isset( $input['social_on_home'])) ? 1 : 0;
	$valid['social_on_pages'] = (isset( $input['social_on_pages'])) ? 1 : 0;
	$valid['social_on_posts'] = (isset( $input['social_on_posts'])) ? 1 : 0;
	$valid['social_on_category'] = (isset( $input['social_on_category'])) ? 1 : 0;
	$valid['social_on_archive'] = (isset( $input['social_on_archive'])) ? 1 : 0;
	$valid['social_before_content'] = (isset( $input['social_before_content'])) ? 1 : 0;
	$valid['social_after_content'] = (isset( $input['social_after_content'])) ? 1 : 0;
	
	return $valid;
}

add_action('admin_menu', 'register_social_ring_admin_menu');
function register_social_ring_admin_menu() {

	$page = add_menu_page(__('Social Ring settings',WP_SOCIAL_RING),__('Social Ring',WP_SOCIAL_RING),'manage_options',WP_SOCIAL_RING,WP_SOCIAL_RING.'_setting_page',plugins_url('/images/av-16x16.png',__FILE__));
	add_action('admin_print_styles-' . $page, 'add_wp_social_ring_css_js');

}

?>