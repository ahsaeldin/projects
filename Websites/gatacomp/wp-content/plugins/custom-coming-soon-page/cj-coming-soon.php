<?php
/*
Plugin Name: CJ Coming Soon
Plugin URI: http://www.cssjockey.com/wordpress-plugins/custom-coming-soon-pages-wordpress-plugin
Description: This plugin shows a 'Custom Coming Soon' page to all users who are not logged in however, the Site Administrators see the fully functional website with the applied theme and active plugins as well as a fully functional Dashboard. Visit our <strong><a href="http://support.cssjockey.com">Support Forum</a></strong> for support, report bugs and request more features and share your themes.
Version: 1.09
Author: CSSJockey
Author URI: http://www.cssjockey.com
/*  Copyright 2009 CSSJockey.com  (email : admin@cssjockey.com)
    This program is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
*/
require_once( dirname(__FILE__) . '../../../../wp-load.php');
ob_start();
if ( !defined('WP_CONTENT_URL') )define( 'WP_CONTENT_URL', get_option('siteurl') . '/wp-content');
if ( !defined('WP_CONTENT_DIR') )define( 'WP_CONTENT_DIR', ABSPATH . 'wp-content' );
/******************************************
 * DEFAULT VARIABLES 
******************************************/
global $cj_splash_path, $cj_splash_shortname, $cj_splash_settings_name, $cj_splash_options;
$cj_splash_plugin_name = "CJ Splash Page";
$cj_splash_path = WP_CONTENT_URL.'/plugins/'.plugin_basename(dirname(__FILE__));
$cj_splash_shortname = strtolower(str_replace(" ", "_", $cj_splash_plugin_name)."_");
$cj_splash_settings_name = $cj_splash_shortname."settings";
/******************************************
 * ADMIN PAGE SETUP - ADMIN SCRIPTS
******************************************/
add_action('admin_init', 'cj_splash_plugin_admin_init');
function cj_splash_plugin_admin_init() {
    global $cj_splash_path;
    wp_register_script('cj_splash_plugin_scripts', $cj_splash_path.'/admin/admin.js');
}
function cj_splash_plugin_admin_styles() {
    wp_enqueue_script('cj_splash_plugin_scripts');
}
function cj_splash_admin_scripts(){
	global $cj_splash_path;
	echo '<link rel="stylesheet" type="text/css" media="screen" href="'.$cj_splash_path.'/admin/admin.css" />';
}
add_action('admin_head', 'cj_splash_admin_scripts');

/******************************************
 * PLUGIN SETTINGS LINK
******************************************/
function cj_splash_plugin_action_links($links, $file){
	static $this_plugin;
	if( ! $this_plugin ) $this_plugin = plugin_basename(__FILE__);
	if( $file == $this_plugin ){
// UPDATE THIS VALUE *****************************************************/
		$settings_link = '<a href="options-general.php?page=coming-soon-settings.php">' . __('Settings') . '</a>';
		$links = array_merge( array($settings_link), $links);
	}
	return $links;
}
add_filter('plugin_action_links', 'cj_splash_plugin_action_links', 10, 2 );

/******************************************
 * ADMIN PAGE SETUP - ADMIN PAGES
******************************************/
add_action('admin_menu', 'cj_splash_admin_menu');
function cj_splash_admin_menu(){
	$page = add_submenu_page('options-general.php', 'CJ Coming Soon', 'CJ Coming Soon', 8, 'coming-soon-settings', 'cj_splash_options_page');
	add_action('admin_print_scripts-' . $page, 'cj_splash_plugin_admin_styles');
}
function cj_splash_options_page(){
	global $cj_splash_settings_name, $cj_splash_shortname;
	$cj_splash_options = array (
	// START EDITING HERE
		    array(  
				"oid" => $cj_splash_shortname."basic_configuration",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Plugin Settings &raquo;",
				"oinfo" => '',
				"otype" => "heading",
                                "ovalue" => 'Basic Settings &raquo;'),
			array(
				"oid" => $cj_splash_shortname."enable_page",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Enable Custom Page",
				"oinfo" => 'You can enable or disable the custom page here.',
				"otype" => "radio",
                                "ovalue" => array("Enable", "Disable")),
                        array(
				"oid" => $cj_splash_shortname."select_theme",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Configure Theme Style",
				"oinfo" => '<b>Default</b>: This will enable the default black theme for the page.<br />
							<b>Custom Background</b>: You can specify desired background color and image below.<br />
							<b>Custom xHTML</b>: You can place your xHTML/CSS/PHP documents in wp-plugins/cj-coming-soon/themes/custom_xHTML/ directory ',
				"otype" => "radio",
                                "ovalue" => array("Default", "Without Timer", "Custom Background", "Custom xHTML")),
			array(
				"oid" => $cj_splash_shortname."login_link",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Login Link in Footer",
				"oinfo" => 'Would you like to display login link in the footer?',
				"otype" => "radio",
                                "ovalue" => array('No', 'Yes')),
			array(
				"oid" => $cj_splash_shortname."logo_url",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Logo Image URL",
				"oinfo" => 'Enter full URL of the logo image. Use <b>transparent .png</b> file for best results, Don\'t worry IE6 is tamed.',
				"otype" => "text",
                                "ovalue" => "yoursite.com/images/logo.png"),
			array(
				"oid" => $cj_splash_shortname."favicon_url",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Favicon URL",
				"oinfo" => 'Enter full URL for your <b>favicon.ico</b> file.',
				"otype" => "text",
                                "ovalue" => "yoursite.com/favicon.ico"),
			array(
				"oid" => $cj_splash_shortname."page_heading",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Page Heading",
				"oinfo" => 'Enter heading in the above box. e.g. Coming Soon! or Under Construction!',
				"otype" => "text",
                                "ovalue" => "Coming Soon!"),
			array(
				"oid" => $cj_splash_shortname."page_msg",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Coming Soon Message",
				"oinfo" => 'Enter a message for the users. Keep it short and <b>Avoid HTML tags.</b>',
				"otype" => "textarea",
                                "ovalue" => "We're not there yet - but getting there - and we really want you to know when we're ready. So be sure to stay updated."),
			array(
				"oid" => $cj_splash_shortname."twitter_username",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Twitter Username",
				"oinfo" => 'Enter your Twitter username (leave it blank if you don\'t wish to display this link )<br />e.g. http://www.twitter.com/<b>cssjockey</b>',
				"otype" => "text",
                                "ovalue" => 'cssjockey'),
			array(
				"oid" => $cj_splash_shortname."facebook_username",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "Facebook Username",
				"oinfo" => 'Enter your facebook username (leave it blank if you don\'t wish to display this link ) <br />e.g. http://www.facebook.com/<b>cssjockey</b>',
				"otype" => "text",
                                "ovalue" => 'cssjockey'),
			array(
				"oid" => $cj_splash_shortname."rss_url",
				"oclass" => $cj_splash_shortname."basic_config",
				"oname" => "RSS URL (if any)",
				"oinfo" => 'Enter RSS URL (leave it blank if you don\'t wish to display this link )<br />e.g. http://feeds.feedburner.com/<b>cssjockey</b>',
				"otype" => "text",
                                "ovalue" => 'http://feeds.feedburner.com/cssjockey'),

			array(
				"oid" => $cj_splash_shortname."custom_bg_settings",
				"oclass" => $cj_splash_shortname."custom_bg_config",
				"oname" => "Custom Background &raquo;",
				"oinfo" => '',
				"otype" => "heading",
                                "ovalue" => "Custom Background &raquo;"),
			array(
				"oid" => $cj_splash_shortname."custom_bg",
				"oclass" => $cj_splash_shortname."custom_bg_config",
				"oname" => "Custom Background Color",
				"oinfo" => 'Enter #HEX value of desired background color. e.g. #000000 for black
							<br /><span style="color:#990000">This will only work with Custom Background Option</span>',
				"otype" => "text",
                                "ovalue" => "#990000"),
			array(
				"oid" => $cj_splash_shortname."custom_bg_img",
				"oclass" => $cj_splash_shortname."custom_bg_config",
				"oname" => "Custom Background Image",
				"oinfo" => 'Enter CSS property with URL and positioning of background image. <br /><b>e.g.</b> url(http://yourwebsite.com/images/bg.jpg) repeat | Leave blank for none.
							<br /><span style="color:#990000">This will only work with Custom Background Option</span>',
				"otype" => "text",
                                "ovalue" => "url(http://yourwebsite.com/images/bg.jpg) repeat"),

		    array(  
				"oid" => $cj_splash_shortname."seo_settings",
				"oclass" => $cj_splash_shortname."head_settings",
				"oname" => "Header Settings &raquo;",
				"oinfo" => '',
				"otype" => "heading",
                                "ovalue" => 'SEO Settings &raquo;'),
		    array(  
				"oid" => $cj_splash_shortname."head_title",
				"oclass" => $cj_splash_shortname."head_settings",
				"oname" => "Document Title",
				"oinfo" => 'Please enter document title, this will show up in Google Search Results once your coming soon page is indexed.',
				"otype" => "text",
                                "ovalue" => 'Coming Soon | Powered by CSSJockey'),
		    array(  
				"oid" => $cj_splash_shortname."head_keywords",
				"oclass" => $cj_splash_shortname."head_settings",
				"oname" => "Meta Keywords",
				"oinfo" => 'Enter some keywords, there\'s no harm adding a few at this stage.',
				"otype" => "text",
                                "ovalue" => 'coming soon page, launch page, under construction page'),
		    array(  
				"oid" => $cj_splash_shortname."head_description",
				"oclass" => $cj_splash_shortname."head_settings",
				"oname" => "Meta Description",
				"oinfo" => 'Enter some text for META Description tag, this will show up in Google search results once your coming soon page is indexed.',
				"otype" => "textarea",
                                "ovalue" => 'Coming Soon Page WordPress Plugin is simply a modern version of the under construction page that you can use if you are about to launch your website, doing some cool enhancements on the design or just fixing some stupid bugs on your WordPress blog or website.'),
		    array(  
				"oid" => $cj_splash_shortname."head_tags",
				"oclass" => $cj_splash_shortname."head_settings",
				"oname" => "Additional Head Tags",
				"oinfo" => 'Enter valid xHTML code to be included within &lt;head&gt;&lt;/head&gt; tags.',
				"otype" => "textarea",
                                "ovalue" => '<!-- Additional Meta Tags -->'),
		    array(  
				"oid" => $cj_splash_shortname."email_settings",
				"oclass" => $cj_splash_shortname."email_settings",
				"oname" => "Email Settings &raquo;",
				"oinfo" => '',
				"otype" => "heading",
                                "ovalue" => 'Email Settings &raquo;'),
			array(
				"oid" => $cj_splash_shortname."email_id",
				"oclass" => $cj_splash_shortname."email_settings",
				"oname" => "Your Email Address",
				"oinfo" => 'Subscription info will be sent to this address.',
				"otype" => "text",
                                "ovalue" => get_option('admin_email')),
			array(
				"oid" => $cj_splash_shortname."email_subject",
				"oclass" => $cj_splash_shortname."email_settings",
				"oname" => "Email Subject Line",
				"oinfo" => 'Enter a subject line to identify and create filters based on the email program you use.',
				"otype" => "text",
                                "ovalue" => "Subscriber's Info @ ".get_bloginfo('name')),
			array(
				"oid" => $cj_splash_shortname."email_thankyou",
				"oclass" => $cj_splash_shortname."email_settings",
				"oname" => "Thank you message",
				"oinfo" => 'Enter a short thank you message to the user once he/she subscribe to the mailing list.',
				"otype" => "textarea",
                                "ovalue" => "Thank you for subscribing to the our Launch Announcement email list."),
		    array(  
				"oid" => $cj_splash_shortname."launch_date",
				"oclass" => $cj_splash_shortname."launch_date",
				"oname" => "Launch Date &raquo;",
				"oinfo" => '',
				"otype" => "heading",
                                "ovalue" => 'Launch Date &raquo;'),
		    array(  
				"oid" => $cj_splash_shortname."launch_day",
				"oclass" => $cj_splash_shortname."launch_date",
				"oname" => "Day of the Month (DD)",
				"oinfo" => 'e.g. 01 - 31',
				"otype" => "text",
                                "ovalue" => '15'),
		    array(  
				"oid" => $cj_splash_shortname."launch_month",
				"oclass" => $cj_splash_shortname."launch_date",
				"oname" => "Month (MM)",
				"oinfo" => 'e.g. 01 - 12',
				"otype" => "text",
                                "ovalue" => '02'),
		    array(  
				"oid" => $cj_splash_shortname."launch_year",
				"oclass" => $cj_splash_shortname."launch_date",
				"oname" => "Year (YYYY)",
				"oinfo" => 'e.g. 2011',
				"otype" => "text",
                                "ovalue" => '2011'),
	// END EDITING HERE
);
/**** INSTALL SETTINGS / ADD OPTIONS ***************************/
foreach($cj_splash_options as $value){
	$saveoptions[$value['oid']] = $value['ovalue'];
}
add_option($cj_splash_settings_name, $saveoptions);
/**** SAVE SETTINGS / UPDATE OPTIONS ***************************/
if(isset($_REQUEST['cjsave'])){
foreach($cj_splash_options as $value){
	$saveoptions[$value['oid']] = $_REQUEST[$value['oid']];
}
update_option($cj_splash_settings_name, $saveoptions);
echo '<div id="message" class="updated fade"><p><strong>Theme settings updated.</strong></p></div>';
}
/**** RESET SETTINGS / DELETE OPTIONS ***************************/
if(isset($_REQUEST['cjreset'])){
delete_option($cj_splash_settings_name);
foreach($cj_splash_options as $value){
	$saveoptions[$value['oid']] = $value['ovalue'];
}
update_option($cj_splash_settings_name, $saveoptions);
echo '<div id="message" class="updated fade"><p><strong>Theme settings updated.</strong></p></div>';
}
/**** SAVE SETTINGS / ADD OPTIONS ***************************/
if(isset($_REQUEST['cjremove'])){
delete_option($cj_splash_settings_name);
$reseturl =  get_bloginfo("wpurl")."/wp-admin/themes.php";
switch_theme('default', 'default');
wp_redirect($reseturl);
}
/**** GET SAVED VALUES ***************************/
function cj_splash_gop($mykey){
global $cj_splash_settings_name;
$sopt = get_option($cj_splash_settings_name);
	foreach($sopt as $key=>$opt){
		if($key == $mykey){
                    if(is_array($opt)){
                        return $opt;
                    }else{
                        return stripcslashes($opt);
                    }
		}
	}
}
?>
<div class="wrap">
<div id="icon-options-general" class="icon32"><br></div>
<h2>Plugin Settings</h2>

<p><b>Upgrade to <a href="http://www.netchillies.com/coming-soon-pages" title="">Pro Version</a> for more themes and options.</b></p>

<div id="cjwarp">
<form action="" method="post">
<?php
foreach($cj_splash_options as $key){ ?>
<?php if($key['otype'] == "heading"){ ?>
	<h1 id="<?php echo $key['oclass']; ?>" class="cjhead"><?php echo $key['ovalue']; ?></h1>
<?php } ?>
<?php if($key['otype'] == "text"){ ?>
<div class="<?php echo $key['oclass']; ?> cjcontent cjclearfix">
	<div class="cjlabel">
		<?php echo $key['oname']; ?>
	</div><!-- /label -->
	<div class="cjfield">
		<input size="35" name="<?php echo $key['oid']; ?>" type="text" value="<?php echo cj_splash_gop($key['oid']); ?>" />
		<span class="cjdesc"><?php echo $key['oinfo']; ?></span>
	</div><!-- /cjfield -->
</div><!-- /cjcontent -->
<?php } ?>
<?php if($key['otype'] == "info"){ ?>
<div class="<?php echo $key['oclass']; ?> cjcontent cjclearfix">
	<div class="cjlabel">
		<label class="cjt-field"><?php echo $key['oname']; ?></label>
	</div><!-- /label -->
	<div class="cjfield">
		<span class="cjt-desc"><?php echo $key['ovalue']; ?></span>
	</div><!-- /cjfield -->
</div><!-- /cjcontent -->
<?php } ?>
<?php if($key['otype'] == "textarea"){ ?>
<div class="<?php echo $key['oclass']; ?> cjcontent cjclearfix">
	<div class="cjlabel">
		<?php echo $key['oname']; ?>
	</div><!-- /label -->
	<div class="cjfield">
		<textarea rows="5" cols="50" name="<?php echo $key['oid']; ?>"><?php echo cj_splash_gop($key['oid']); ?></textarea>
		<span class="cjdesc"><?php echo $key['oinfo']; ?></span>
	</div><!-- /cjfield -->
</div><!-- /cjcontent -->
<?php } ?>
<?php if($key['otype'] == "select"){ ?>
<div class="<?php echo $key['oclass']; ?> cjcontent cjclearfix">
	<div class="cjlabel">
		<?php echo $key['oname']; ?>
	</div><!-- /label -->
	<div class="cjfield">
		<?php $soptions = $key['ovalue']; ?>
		<select class="cjselect" name="<?php echo $key['oid']; ?>">
			<option value="Please Select">Please Select</option>
			<?php
			foreach($soptions as $svalue){ ?>
				<option <?php if($svalue == cj_splash_gop($key['oid'])){echo 'selected="selected"';} ?> value="<?php echo $svalue ?>"><?php echo $svalue ?></option>
			<?php }
			?>
		</select>
		<span class="cjdesc"><?php echo $key['oinfo']; ?></span>
	</div><!-- /cjfield -->
</div><!-- /cjcontent -->
<?php } ?>
<?php if($key['otype'] == "radio"){ ?>
<div class="<?php echo $key['oclass']; ?> cjcontent cjclearfix">
	<div class="cjlabel">
		<?php echo $key['oname']; ?>
	</div><!-- /label -->
	<div class="cjfield">
		<?php $roptions = $key['ovalue']; ?>
			<?php
			foreach($roptions as $rvalue){ ?>
				<span class="cjradio">
				<input type="radio" class="cjradio" name="<?php echo $key['oid']; ?>" <?php if($rvalue == cj_splash_gop($key['oid'])){echo 'checked="checked"';} ?> value="<?php echo $rvalue; ?>" /> <?php echo $rvalue; ?>
				</span>
			<?php }
			?>
		<span class="cjdesc"><?php echo $key['oinfo']; ?></span>
	</div><!-- /cjfield -->
</div><!-- /cjcontent -->
<?php } ?>
<?php if($key['otype'] == "categories"){ ?>
<div class="<?php echo $key['oclass']; ?> cjcontent cjclearfix">
	<div class="cjlabel">
		<?php echo $key['oname']; ?>
	</div><!-- /label -->
	<div class="cjfield">
		<input size="35" name="<?php echo $key['oid']; ?>" type="text" value="<?php echo cj_splash_gop($key['oid']); ?>" /> <a id="cjshowcatids" href="#" title="Display Category IDs">[Show Category IDs]</a>
		<span class="cjdesc"><?php echo $key['oinfo']; ?></span>
		<div>
	<ul class="cjids cjcats cjclearfix">
		<?php
		    $category_ids = get_all_category_ids();
			foreach($category_ids as $cat_id){
				$cat_name = get_cat_name($cat_id);
				echo '<li><b>' .$cat_id . '</b> - '.$cat_name.'</li>';
			}
		?>
	</ul>
		</div>
	</div><!-- /cjfield -->
</div><!-- /cjcontent -->
<?php } ?>
<?php if($key['otype'] == "pages"){ ?>
<div class="<?php echo $key['oclass']; ?> cjcontent cjclearfix">
	<div class="cjlabel">
		<?php echo $key['oname']; ?>
	</div><!-- /label -->
	<div class="cjfield">
		<input size="35" name="<?php echo $key['oid']; ?>" type="text" value="<?php echo cj_splash_gop($key['oid']); ?>" /> <a id="cjshowpageids" href="#" title="Display Page IDs">[Show Page IDs]</a>
		<span class="cjdesc"><?php echo $key['oinfo']; ?></span>
		<div>
	<ul class="cjids cjpages cjclearfix">
		<?php
		    $pages = get_pages(); 
				foreach ($pages as $pagg) {
					echo '<li><b>' .$pagg->ID . '</b> - '.$pagg->post_title.'</li>';
				}
		?>
	</ul>
		</div>
	</div><!-- /cjfield -->
</div><!-- /cjcontent -->
<?php } 
} ?>
<div class="cjbuttons cjclearfix">
	<input name="cjsave" class="button-primary" type="submit" value="Save Settings" />
	<input name="cjreset" class="button" type="submit" value="Restore Defaults" />
</div>
<div style="border-top:0px;" class="cjbuttons cjclearfix"> 
If you like this plugin, consider
<a target="_blank" href="http://www.cssjockey.com/files/donate.php" title="Donate">making a donation</a>.
Your support will help me spend more time on further development of this plugin and keep it free forever.
</div>
</form>
<div style="border-top:0px;" class="cjbuttons cjclearfix">
	<b>New @ CSSJockey.com</b> &raquo;
<?php // Get RSS Feed(s)
include_once(ABSPATH . WPINC . '/feed.php');
$rss = fetch_feed('http://feeds.feedburner.com/cssjockey');
$maxitems = $rss->get_item_quantity(1); 
$rss_items = $rss->get_items(0, $maxitems); 
?>
    <?php if ($maxitems == 0) echo '<li>No items.</li>';
    else
    // Loop through each feed item and display each item as a hyperlink.
    foreach ( $rss_items as $item ) : ?>
        <a target="_blank" href='<?php echo $item->get_permalink(); ?>' title='<?php echo $item->get_title(); ?>'>
        	<?php echo $item->get_title(); ?>
		</a>
    <?php endforeach; ?>
</div><!-- / -->
<iframe src="http://www.cssjockey.com/files/theme-admin.php" scrolling="no" frameborder="0" width="100%"></iframe>
</div><!-- /cjt-wrap -->
</div><!-- /wrap -->	
<?php
    //PLUGIN SETTINGS CHECK
    if(get_option('sp_settings_check') == 0) {
        update_option('sp_settings_check' , '1');
    }
} //options page
register_deactivation_hook(__FILE__, 'cj_splash_deactivate');
function cj_splash_deactivate(){
    global $cj_splash_settings_name;
    delete_option($cj_splash_settings_name);
}
/******************************************
 * PLUGIN SETTINGS CHECK
******************************************/
register_activation_hook(__FILE__, 'sp_settings_check');
function sp_settings_check() {
    update_option('sp_settings_check', '0');
}//continued above // options page
if(get_option('sp_settings_check') == '0') {
    header('location:'.get_bloginfo('wpurl').'/wp-admin/options-general.php?page=coming-soon-settings'); // Change URL as per page
}
/************************************************************************************************************/
/******************************************* ADMIN SETUP ENDS HERE ******************************************/
/************************************************************************************************************/

/******************************************
 * PLUGIN FUNCTIONS
******************************************/
add_action('get_header', 'sp_override');
function cj_splash_top($mykey){
global $cj_splash_settings_name, $cj_splash_shortname;
$mykey = $cj_splash_shortname.$mykey;
$sopt = get_option($cj_splash_settings_name);
foreach($sopt as $key=>$opt) {
    if($key == $mykey) {
        return stripcslashes($opt);
    }
}

}

function sp_override(){
    $sp_style = cj_splash_top("select_theme");
    $enable = cj_splash_top('enable_page');
    if($enable == "Enable"){
        if(!is_user_logged_in()) {
            switch ($sp_style) {
                case "Default": include('themes/default/index.php'); break;
                case "Custom Background": include('themes/custom_bg/index.php'); break;
                case "Custom xHTML": include('themes/custom_xhtml/index.php'); break;
                case "Without Timer": include('themes/default_no_timer/index.php'); break;
                default: include('themes/default/index.php');
            }
        }
        elseif (!current_user_can('level_10')) {
        }
    }
}
?>