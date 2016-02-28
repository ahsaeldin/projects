<?php
/*
Plugin Name: Custom Admin Branding
Plugin URI: http://redlettersstudio.com/wordpress-custom-admin-branding/
Description: Allows you to brand your WordPress install for clients.  Display custom images and styles for the login screen, admin header and footer.
Author: Josh Byers
Version: 2.0
Author URI: http://www.redlettersstudio.com
*/ 

add_action('admin_menu', 'mt_add_pages');
add_action('admin_init', 'admin_branding_options_init' );


// Init plugin options to white list our options
function admin_branding_options_init(){
	register_setting( 'admin_branding_options', 'admin_branding_link' );	
}

// Add menu page
function mt_add_pages() {

	$custombranding_admin = add_options_page('Custom Admin Branding', 'Custom Admin Branding', 8, 'brandingoptions', 'admin_branding_options_page');

	
//add help instructions for WordPress sliding help menu
$admin_branding_help = "
<style>
h6 {
	font-size:0.85em;
	border-bottom:1px solid #d1d1d1;
	padding-bottom:3px;
	margin:2.33em 0 10px;}

ul.help_list {
	margin-top:10px;}
	
ul.help_list li {
	list-style-type:disc;
	margin-left:20px;}
</style>

<p>Well, you obviously have the plugin activated so thats out of the way.  Now on to the customization of your WordPress Admin.</p>
<p>Nearly every visual feature in the login screen, header section, and footer section can be modified to fit your custom color and image scheme.</p>

<h6>Changing Colors</h6>
<p>The options that support color change have a swatch next to the form field.  You can either click the color swatch or click into the form field.  Doing so will bring up a palette of color options.  You are not restricited by the choices in the palette.  You can enter any 6 figure hexidecimal color code.  You must however precede your code with the pound sign (#).</p>

<h6>Customizing Images</h6>
There are five images that you can customize:
<ul class='help_list'>
	<li>Login Form Background</li>
	<li>Login Form Login Button Background</li>
	<li>Admin Header Logo</li>
	<li>Favorite Actions Button Background</li>
	<li>Footer Logo</li>
</ul>

<p>All of the default images are in the images folder which is in the custom-admin-branding plugin folder.
To change an image:</p>
<ol>
	<li>Create your custom images with the provided Photoshop templates. (They are in the psd_image_templates folder)</li>
	<li>Upload your custom image(s) to the images folder.</li>
	<li>Enter the file name of the image you wish to use in the proper form field in the options page.</li>
	<li>Click the save changes button at the bottom of options page and you will see your changes immediately.</li>
</ol>

<p>*********Important***********</p>
To avoid having your custom images overwritten when you upgrade the plugin, you'll need to name them something different than the default names which are:
<ul class='help_list'>
	<li>wordpress_blue.jpg</li>
	<li>button-grad.png</li>
	<li>logo_wordpress_gray.png</li>
	<li>fav.png</li>
	<li>footer_wordpress_gray.png</li>
</ul>


<h6>Advanced Header Customization</h6>
<p>The form field for the Admin Header Background Color will accept a full css properties for the background selector.  So if you wanted to have a repeating image for your admin header you would enter in:<br />
<blockquote>transparent url(http://yoursitename.com/wp-content/plugins/custom-admin-branding/images/header_bg_wordpress_gray.png) repeat-x</blockquote></p>

<h6>Changing the Admin Footer Link</h6>
<p>You can have the image in the footer link back to your site, another site, or no site if you wish.  Just fill in full url in the Footer Link form field.</p>

"
;
add_contextual_help($custombranding_admin, $admin_branding_help);
}
function on_load_page() {
		//ensure, that the needed javascripts been loaded to allow drag/drop, expand/collapse and hide/show of boxes
		wp_enqueue_script('common');
		wp_enqueue_script('wp-lists');
		wp_enqueue_script('postbox');
}
// admin_branding_options_page() displays the page content for the Admin Branding Options submenu
function admin_branding_options_page() {
$options = get_option('admin_branding_link');
	// default options to be loaded on plugin activation
	$options['top_page_border_show'] = 'normal';
	$options['top_page_border_color'] = '#464646';
	$options['back_to_blog_color'] = '#cccccc';
	$options['back_to_blog_hover_color'] = '#ffffff';
	$options['login_form_bg'] = 'login_wordpress_gray.jpg';
	$options['login_form_label'] = '#999999';
	$options['input_text_color'] = '#555555';
	$options['footer_link'] = 'http://wordpress.org';  
	$options['login_button_color'] = '#21759B';
	$options['login_button_background'] = 'button-grad.png';
	$options['login_border_color'] = '#298CBA';
	$options['login_border_hover_color'] = '#13455b';
	$options['login_text_color'] = '#eaf2a';
	$options['admin_header_bg_img'] = '';
	$options['admin_header_bg_clr'] = '#464646';
	$options['admin_header_logo'] = 'logo_wordpress_gray.png';
	$options['fav_actions_btn_color'] = '#797979';
	$options['fav_actions_btn_background'] = 'fav.png';
	$options['fav_actions_bdr'] = '#777777';
	$options['user_info_txt'] = '#999999';
	$options['user_info_lnk'] = '#cccccc';
	$options['user_info_lnk_hvr'] = '#ffffff';
	$options['admin_footer_logo'] = 'footer_wordpress_gray.png';
	$options['lost_password_color'] = '#21759B';
	$options['lost_password_hover_color'] = '#D54E21';
	add_option( 'admin_branding_link', $options );
 ?>
 
<?php 
// calls and variables to be loaded in WordPress admin for Live Login Preview
$options = get_option('admin_branding_link');
	$top_page_border_show = $options['top_page_border_show'];
	$top_page_border_color = $options['top_page_border_color'];
	$back_to_blog_color = $options['back_to_blog_color'];
	$back_to_blog_hover_color = $options['back_to_blog_hover_color'];
	$login_form_bg = $options['login_form_bg'];
	$login_form_label = $options['login_form_label'];
	$input_text_color = $options['input_text_color'];
	$login_button_color = $options['login_button_color'];
	$login_button_background = $options['login_button_background'];
	$login_border_color = $options['login_border_color'];
	$login_border_hover_color = $options['login_border_hover_color'];
	$login_text_color = $options['login_text_color'];
	$lost_password_color = $options['lost_password_color'];
	$lost_password_hover_color = $options['lost_password_hover_color'];
?>
<style >
#ColorPickerDiv {
    display: block;
    display: none;
    position: relative;
    border: 1px solid #777;
    background: #fff}
#ColorPickerDiv TD.color {
	cursor: pointer;
	font-size: xx-small;
	font-family: 'Arial' , 'Microsoft Sans Serif';}
#ColorPickerDiv TD.color label{
	cursor: pointer;}
.ColorPickerDivSample {
	margin: 0px 0px 0px 4px;
	border: solid 1px #000;
	padding: 0px 10px;	
	position: relative;
	cursor: pointer;}
input.wide {
	width:308px;}
</style>
<script type="text/javascript" src="<?php echo $site_url; ?>/wp-content/plugins/custom-admin-branding/js/ColorPicker.js"></script>
<script type="text/javascript">
jQuery(function($) {
   $(".vtrColorPicker").attachColorPicker();
});
</script>

<div id="howto-metaboxes-general" class="wrap">
	<?php screen_icon('options-general'); ?>
    <h2>Custom Admin Branding Options</h2>
<div class="left_side" style="width:540px;float:left;margin-right:20px;">
	<div class="metabox-holder" >       
        <div class="postbox">
            <h3 class="hndle"><span>Login Screen Options</span></h3>
            <div class="inside">
                <form method="post" action="options.php">
                <?php settings_fields('admin_branding_options');?>
                <?php $options = get_option('admin_branding_link');?>
                <? $site_url = get_settings('siteurl'); ?>
                <table class="form-table">
                    <tr valign="top">
                        <th scope="row">Top Page Border</th>       
                            <td>Show: <input name="admin_branding_link[top_page_border_show]" type="radio" value="normal" <?php checked('normal', $options['top_page_border_show']); ?> />
                Hide: <input name="admin_branding_link[top_page_border_show]" type="radio" value="none" <?php checked('none', $options['top_page_border_show']); ?> />
                            <br />Color:<br /><input type="text" class="vtrColorPicker" name="admin_branding_link[top_page_border_color]" value="<?php echo $options['top_page_border_color']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Back to Blog Link</th>       
                            <td>Link Color:<br /><input type="text" class="vtrColorPicker" name="admin_branding_link[back_to_blog_color]" value="<?php echo $options['back_to_blog_color']; ?>" /><br />Hover Color:<br /><input type="text" class="vtrColorPicker" name="admin_branding_link[back_to_blog_hover_color]" value="<?php echo $options['back_to_blog_hover_color']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Form Background</th>
                            <td><input type="text" class="wide" name="admin_branding_link[login_form_bg]" value="<?php echo $options['login_form_bg']; ?>" />
                                    </td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Input Label Color</th>
                            <td><input class="vtrColorPicker" type="text" name="admin_branding_link[login_form_label]" value="<?php echo $options['login_form_label']; ?>" /></td>
                    </tr>  	
                    <tr valign="top">
                        <th scope="row">Input Text Color</th>
                            <td><input class="vtrColorPicker" type="text" name="admin_branding_link[input_text_color]" value="<?php echo $options['input_text_color']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Login Button Background</th>
                            <td><input type="text" class="wide" name="admin_branding_link[login_button_background]" value="<?php echo $options['login_button_background']; ?>" />
                            </td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Login Button Border Color</th>
                            <td><input type="text" class="vtrColorPicker" name="admin_branding_link[login_button_color]" value="<?php echo $options['login_button_color']; ?>" />
                            </td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Login Button Hover Color</th>
                            <td><input type="text" class="vtrColorPicker" name="admin_branding_link[login_border_hover_color]" value="<?php echo $options['login_border_hover_color']; ?>" />
                            </td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Login Button Text Color</th>
                            <td><input type="text" class="vtrColorPicker" name="admin_branding_link[login_text_color]" value="<?php echo $options['login_text_color']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Lost Your Password Link</th>       
                            <td>Link Color:<br /><input type="text" class="vtrColorPicker" name="admin_branding_link[lost_password_color]" value="<?php echo $options['lost_password_color']; ?>" /><br />Hover Color:<br /><input type="text" class="vtrColorPicker" name="admin_branding_link[lost_password_hover_color]" value="<?php echo $options['lost_password_hover_color']; ?>" /></td>
                    </tr>
                </table>
            </div>
        </div>        
     </div>
     
     <div class="metabox-holder">       
        <div class="postbox">
            <h3 class="hndle"><span>Admin Header Options</span></h3>
            <div class="inside">
                <table class="form-table">
                    <tr valign="top">
                        <th scope="row">Admin Header Logo</th>
                            <td><input type="text" class="wide" name="admin_branding_link[admin_header_logo]" value="<?php echo $options['admin_header_logo']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Admin Header Background Color</th>
                            <td><input type="text" class="vtrColorPicker" name="admin_branding_link[admin_header_bg_clr]" value="<?php echo $options['admin_header_bg_clr']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Favorite Actions Button Color</th>
                            <td><input type="text" class="vtrColorPicker" name="admin_branding_link[fav_actions_btn_color]" value="<?php echo $options['fav_actions_btn_color']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Favorite Actions Button Background</th>
                            <td><input type="text" class="wide" name="admin_branding_link[fav_actions_btn_background]" value="<?php echo $options['fav_actions_btn_background']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">Favorite Actions Button Border Color</th>
                            <td><input type="text" class="vtrColorPicker" name="admin_branding_link[fav_actions_bdr]" value="<?php echo $options['fav_actions_bdr']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">User Info Text Color</th>
                            <td><input type="text" class="vtrColorPicker" name="admin_branding_link[user_info_txt]" value="<?php echo $options['user_info_txt']; ?>" /></td>
                    </tr>
                    <tr valign="top">
                        <th scope="row">User Info Link Color</th>
                            <td>Link Color:<br /><input class="vtrColorPicker" type="text" name="admin_branding_link[user_info_lnk]" value="<?php echo $options['user_info_lnk']; ?>" /><br />Hover Color:<br /><input class="vtrColorPicker" type="text" name="admin_branding_link[user_info_lnk_hvr]" value="<?php echo $options['user_info_lnk_hvr']; ?>" /></td>
                    </tr>
            	</table>
            </div> 
        </div>       
     </div>
     
     <div class="metabox-holder">       
        <div class="postbox">
            <h3 class="hndle"><span>Admin Footer Options</span></h3>
            <div class="inside">
                <table class="form-table">
                <tr valign="top">
                    <th scope="row">Admin Footer Logo</th>
                        <td><input type="text" class="wide" name="admin_branding_link[admin_footer_logo]" value="<?php echo $options['admin_footer_logo']; ?>" /></td>
                </tr>
                <tr valign="top">
                    <th scope="row">Footer Link <br /><em>(use full url)</em></th>
                        <td><input type="text" class="wide" name="admin_branding_link[footer_link]" value="<?php echo $options['footer_link']; ?>" /></td>
                </tr>
            </table>
            </div>
        </div>       
     </div>
     
     <p class="submit">
		<input type="submit" class="button-primary" value="<?php _e('Save Changes') ?>" />
		</p>
		</form>
</div>
</div><!-- end .left_side -->
<div class="right_side" style="float:left;width:362px;"
<div class="metabox-holder">       
        <div class="postbox">
            <h3 class="hndle"><span>Login Screen Preview</span></h3>
            <div class="inside" style="padding:20px;">
<style>
p.backtoblog {
	background-color:<?php echo $top_page_border_color; ?>;
	height:20px;
	display:<?php echo $top_page_border_show; ?>;
	-moz-border-radius-bottomleft:6px;
	-moz-border-radius-bottomright:6px;
	-moz-border-radius-topleft:6px;
	-moz-border-radius-topright:6px;
	width:310px;
	margin-left:5px;}
p.backtoblog a {
	color:<?php echo $back_to_blog_color; ?>;
	padding:1px 0 0 11px;
	display:block;
	text-decoration:none;}
p.backtoblog a:hover {
	color:<?php echo $back_to_blog_hover_color; ?>;}

#loginform {
	padding-top:90px;
	background:transparent url(<?php get_settings('siteurl'); ?>/wp-content/plugins/custom-admin-branding/images/<?php echo $login_form_bg; ?>) center top no-repeat;}
.loginform {
	padding: 16px 16px 59px}
#loginform p {
	margin-bottom: 0;}

.form_lables {
	font-size:13px;
	font-family:'Lucida Grande',Verdana,Arial,'Bitstream Vera Sans',sans-serif;
	color:<?php echo $login_form_label; ?>;}

#loginform input.input {
	border:1px solid #E5E5E5;
	font-size:24px; 
	width:97%;
	margin:0;
	color:<?php echo $input_text_color; ?>;}

p.forgetmenot {
	float:left;
	line-height:19px;
	padding-top:14px;}
	
	p.forgetmenot .form_lables {
		font-size:11px;}

p.submit preview {
	float:right;
	padding:12px 0;}

#loginform p.submit input {
	background:<?php echo $login_button_color; ?> url(<?php get_settings('siteurl'); ?>/wp-content/plugins/custom-admin-branding/images/<?php echo $login_button_background; ?> ) repeat-x scroll left top;
	color:<?php echo $login_text_color; ?> !important;
	border-color:<?php echo $login_border_color; ?>; }

#loginform p.submit input:hover {
	border-color:<?php echo $login_border_hover_color; ?>;
	color:<?php echo $login_text_color; ?> !important;}

p.lost_password {
	font-size:11px;
	padding-left:10px;}
	
p.lost_password a {
	color:<?php echo $lost_password_color; ?>;}

p.lost_password a:hover {
	color:<?php echo $lost_password_hover_color; ?>;}
</style>
	<div id="login" style="width:320px;">

		<p class="backtoblog"><a href="http://redlettersstudio.com/" title="Are you lost?">&larr; Back to <?php echo get_settings('blogname'); ?></a></p>
	<form name="loginform" id="loginform" class="loginform" action="http://redlettersstudio.com/wp-login.php" method="post">
		<p>
			<label class="form_lables">Username<br />
			<input type="text" name="log" id="user_login" class="input" value="admin" size="20" tabindex="10" /></label>
		</p>
		<p>
			<label class="form_lables">Password<br />
			<input type="password" name="pwd" id="user_pass" class="input" value="" size="20" tabindex="20" /></label>
		</p>
		<p class="forgetmenot"><label class="form_lables"><input name="rememberme" type="checkbox" id="rememberme" value="forever" tabindex="90" /> Remember Me</label></p>
		<p class="submit preview">
			<input type="submit" name="wp-submit" id="wp-submit" class="button-primary" value="Log In" tabindex="100" />
			<input type="hidden" name="redirect_to" value="http://redlettersstudio.com/wp-admin/" />
			<input type="hidden" name="testcookie" value="1" />
		</p>
	</form>
		<p id="nav" class="lost_password">
			<a href="http://redlettersstudio.com/wp-login.php?action=lostpassword" title="Password Lost and Found">Lost your password?</a>
		</p>
	</div><!-- end #login -->
</div>
</div>
<!--<div class="metabox-holder">       
        <div class="postbox">
            <h3 class="hndle"><span>Please Donate</span></h3>
            <div class="inside" style="padding:10px;">
            	<p>If you have found this plugin useful please consider donating to keep the development and support current.  Thanks.
<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
<input type="hidden" name="cmd" value="_s-xclick">
<input type="hidden" name="hosted_button_id" value="2058369">
<input type="image" src="https://www.paypal.com/en_US/i/btn/btn_donate_LG.gif" border="0" name="submit" alt="">
<img alt="" border="0" src="https://www.paypal.com/en_US/i/scr/pixel.gif" width="1" height="1">
</form></p>
</div>
</div>
</div> -->
</div>
</div>

<?php
 
}

/*This is the function that displays the custom login screen.*/
function custom_login() {
$options = get_option('admin_branding_link');
	$top_page_border_show = $options['top_page_border_show'];
	$top_page_border_color = $options['top_page_border_color'];
	$back_to_blog_color = $options['back_to_blog_color'];
	$back_to_blog_hover_color = $options['back_to_blog_hover_color'];
	$login_form_bg = $options['login_form_bg'];
	$login_form_label = $options['login_form_label'];
	$input_text_color = $options['input_text_color'];  
	$login_button_color = $options['login_button_color'];
	$login_button_background = $options['login_button_background'];
	$login_border_color = $options['login_border_color'];
	$login_border_hover_color = $options['login_border_hover_color'];
	$login_text_color = $options['login_text_color'];
	$lost_password_color = $options['lost_password_color'];
	$lost_password_hover_color = $options['lost_password_hover_color'];

echo '
<style>
/* Diplays the custom graphics for the login screen*/
p#backtoblog {
	display:'.$top_page_border_show.';}
	
.login #backtoblog a {
	color:'.$back_to_blog_color.'}
	
.login #backtoblog a:hover {
	color:'.$back_to_blog_hover_color.'}
		
body.login {
	border:'.$top_page_border_show.';
	border-top-color:'.$top_page_border_color.';}
	
#login form {
	background:transparent url(' . get_settings('siteurl') . '/wp-content/plugins/custom-admin-branding/images/'.$login_form_bg.' ) center top no-repeat;
	padding-top:100px;}

label {
	color:'.$login_form_label.'}

input {
	color:'.$input_text_color.'}

input.button-primary {
	background:'.$login_button_color.' url(' . get_settings('siteurl') . '/wp-content/plugins/custom-admin-branding/images/'.$login_button_background.' ) repeat-x scroll left top;
	color:'.$login_text_color.' !important;}

p.submit input:hover {
	border-color:'.$login_border_hover_color.';}

.login #nav a {
	color:'.$lost_password_color.' !important;}
	
.login #nav a:hover {
	color:'.$lost_password_hover_color.' !important;}

/* Hides the default Wordpress Login content*/	
#login form {
	-moz-border-radius: 11px;
	-khtml-border-radius: 0;
	-webkit-border-radius: 0;
	-moz-box-shadow:none;
	-webkit-box-shadow: none;
	-khtml-box-shadow: none;
	box-shadow: none;
	border-radius: 0px;
	border: none;}

#login h1 {
	display:none;}

</style>
	';
}

add_action('login_head', 'custom_login');

/*This function places the custom header graphic at the top of every WordPress Admin page.*/
function custom_header() {
$options = get_option('admin_branding_link');
	$admin_header_bg_img = $options['admin_header_bg_img'];
	$admin_header_bg_clr = $options['admin_header_bg_clr'];
	$admin_header_logo = $options['admin_header_logo'];
	$fav_actions_btn_color = $options['fav_actions_btn_color'];
	$fav_actions_btn_background = $options['fav_actions_btn_background'];
	$fav_actions_bdr = $options['fav_actions_bdr'];
	$user_info_txt = $options['user_info_txt'];
	$user_info_lnk = $options['user_info_lnk'];
	$user_info_lnk_hvr = $options['user_info_lnk_hvr'];
	
echo '
<style>
/* Hides the default WordPress Admin header content*/	
#wphead  {
	display:none;}
	
/*Custom classes for the WordPress Admin header*/		
#wphead_custom {
	background:'.$admin_header_bg_clr.';
	height:100px;}
			
#wphead-info {
	padding-top:27px;}

#header-logo_custom {
	float: left;}

#favorite-first {
	background:'.$fav_actions_btn_color.' url(' . get_settings('siteurl') . '/wp-content/plugins/custom-admin-branding/images/'.$fav_actions_btn_background.' ) repeat-x scroll left center;
	border-color:'.$fav_actions_bdr.'}

#user_info p {
	color:'.$user_info_txt.';}

#user_info a:link {
	color:'.$user_info_lnk.';}

#user_info a:hover {
	color:'.$user_info_lnk_hvr.';}
</style>
	
<div id="wphead_custom">
	<a href="' . get_settings('siteurl') . '"><img id="header-logo_custom" src="' . get_settings('siteurl') . '/wp-content/plugins/custom-admin-branding/images/'.$admin_header_logo.'"</a>
	<div id="wphead-info">
	 	<div id="user_info">
	 		<p>Howdy, <a href="'. get_settings('siteurl') . '/wp-admin/profile.php">'; 	 	
	 			global $current_user;
				if (isset($current_user->user_firstname)){
				get_currentuserinfo(); echo($current_user->user_firstname);
				} else {
				get_currentuserinfo(); echo($current_user->user_login);
				} 
	 			echo '</a> | <a href="'. get_settings('siteurl') . '/wp-admin/tools.php">Turbo</a> | <a href="'. wp_logout_url().'">Log Out</a> 
			</p>
		</div>';
	 	
	 	echo favorite_actions();
	 	echo '
	</div>
</div>';

}



add_action('admin_head', 'custom_header', 11);

/*This function places the custom footer at the bottom of every WordPress Admin page.  Change the file in the images folder to replace.  You will also want to change the link (in the code below) that the footer image is pointing to.*/

function custom_footer() {
$options = get_option('admin_branding_link');
	$footer_link = $options["footer_link"];
	$admin_footer_logo = $options['admin_footer_logo'];
	
echo '
<style>
/* Hides the default WordPress Admin footer content*/	
#footer {
	display:none;}
		
/*Custom classes for the WordPress Admin footer*/	
#footer_custom {
	clear: both;
	text-align: center;
	width: 500px;
	margin: auto;
	height: 100px;}

#footer_custom .docs {
	padding-top: 0px;
	line-height: 100%;}

#footer_image {
	border:none;}
</style>
	
	
<div id="footer_custom">	
	<p><a href="'.$footer_link.'" id="footer_image"><img src="' . get_settings('siteurl') . '/wp-content/plugins/custom-admin-branding/images/'.$admin_footer_logo.'" alt="'.$footer_link.'" /></a>
	</p>
<p class="docs"> Admin Version '. get_bloginfo('version') . '</p>
</div>';
}

add_action('admin_footer', 'custom_footer');
?>