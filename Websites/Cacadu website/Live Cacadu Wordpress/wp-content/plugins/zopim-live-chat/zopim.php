<?php

/*
Plugin Name: Zopim Widget
Plugin URI: http://www.zopim.com/?iref=wp_plugin
Description: Zopim embeds a chatbar on your website, so that any visitor can chat with you directly by clicking on the chatbar.
Author: Zopim
Version: 1.2.0
Author URI: http://www.zopim.com/?iref=wp_plugin
*/

define('ZOPIM_SCRIPT_DOMAIN',         "zopim.com");
define('ZOPIM_BASE_URL',              "https://www.zopim.com/");
define('ZOPIM_SIGNUP_REDIRECT_URL',   ZOPIM_BASE_URL."?iref=wordpress_plugin#signup");
define('ZOPIM_GETACCOUNTDETAILS_URL', ZOPIM_BASE_URL."plugins/getAccountDetails");
define('ZOPIM_SETDISPLAYNAME_URL',    ZOPIM_BASE_URL."plugins/setDisplayName");
define('ZOPIM_IMINFO_URL',            ZOPIM_BASE_URL."plugins/getImSetupInfo");
define('ZOPIM_IMREMOVE_URL',          ZOPIM_BASE_URL."plugins/removeImSetup");
define('ZOPIM_LOGIN_URL',             ZOPIM_BASE_URL."plugins/login");
define('ZOPIM_SIGNUP_URL',            ZOPIM_BASE_URL."plugins/createTrialAccount");
define('ZOPIM_THEMES_LIST',           "http://zopim.com/assets/dashboard/themes/window/plugins-themes.txt");
define('ZOPIM_COLORS_LIST',           "http://zopim.com/assets/dashboard/themes/window/plugins-colors.txt");
define('ZOPIM_LANGUAGES_URL',         "http://translate.zopim.com/projects/zopim/");
define('ZOPIM_DASHBOARD_URL',         "http://dashboard.zopim.com/?iref=wp_plugin");
define('ZOPIM_SMALL_LOGO',            "http://zopim.com/assets/branding/zopim.com/chatman/online.png");
define('ZOPIM_IM_LOGOS',              "http://www.zopim.com/static/images/im/");
define('ZOPIM_THEMES_URL',            "http://");
define('ZOPIM_COLOURS_URL',           "http://");

require_once dirname( __FILE__ ) . '/accountconfig.php';
require_once dirname( __FILE__ ) . '/customizewidget.php';
require_once dirname( __FILE__ ) . '/imintegration.php';

// We need some CSS to position the paragraph
function zopimme() {
  global $current_user, $zopimshown;

  $code = get_option('zopimCode');

  if (($code == "" || $code=="zopim") && (!ereg("zopim", $_GET["page"]))&& (!ereg("zopim", $_SERVER["SERVER_NAME"]))) { return; }

  // dont show this more than once
  if (isset($zopimshown) && $zopimshown == 1) { return; }
  $zopimshown = 1;

  echo "<!--Start of Zopim Live Chat Script-->
<script type=\"text/javascript\">
window.\$zopim||(function(d,s){var z=\$zopim=function(c){z._.push(c)},$=z.s=
d.createElement(s),e=d.getElementsByTagName(s)[0];z.set=function(o){z.set.
_.push(o)};z._=[];z.set._=[];$.async=!0;$.setAttribute('charset','utf-8');
$.src='//cdn.zopim.com/?".$code."';z.t=+new Date;$.
type='text/javascript';e.parentNode.insertBefore($,e)})(document,'script');
</script>
<!--End of Zopim Live Chat Script-->";

  $theoptions = array();
  if (get_option('zopimLang') != "" && get_option('zopimLang') != "--")
	 $theoptions[] = " language: '".get_option('zopimLang')."'";

  if (isset($current_user) && get_option("zopimGetVisitorInfo") == "checked") {
	 $ul = $current_user->data->first_name;
	 $useremail = $current_user->data->user_email;
	 if ($ul!="" && $useremail != "")
		$theoptions[] = "name: '$ul', email: '$useremail'";
  }

  echo "\n<script type=\"text/javascript\">\n\$zopim( function() {";

  if (count($theoptions) > 0)
	 echo '$zopim.livechat.set({'.implode(", ", $theoptions)."});";

  get_option('zopimPosition')    != "" && print("\n\$zopim.livechat.button.setPosition('".get_option('zopimPosition')."');");
  get_option('zopimTheme')       != "" && print("\n\$zopim.livechat.window.setTheme('".get_option('zopimTheme')."');");
  get_option('zopimColor')       != "" && print("\n\$zopim.livechat.window.setColor('".get_option('zopimColor')."');");

  if (get_option('zopimUseGreetings') == "zopimUseGreetings") {
	  if (get_option('zopimGreetings') != "") {
			$greetings = json_to_array(get_option('zopimGreetings'));
			foreach ($greetings as $i => $v) {
			 foreach ($v as $j => $k) {
				$greetings->$i->$j = str_replace("\r\n", "\\n", $greetings->$i->$j);
			  }
		  }
			echo "\n\$zopim.livechat.setGreetings({
'online' : ['".addslashes($greetings->online->bar)."', '".addslashes($greetings->online->window)."'],
'offline': ['".addslashes($greetings->offline->bar)."', '".addslashes($greetings->offline->window)."'],
'away'   : ['".addslashes($greetings->away->bar)."', '".addslashes($greetings->away->window)."']  });";
	  }
  }

	if (get_option('zopimUseBubble') == "zopimUseBubble") {
	  get_option('zopimBubbleTitle') != "" && print("\n\$zopim.livechat.bubble.setTitle('".addslashes(get_option('zopimBubbleTitle'))."');");
	  get_option('zopimBubbleText')  != "" && print("\n\$zopim.livechat.bubble.setText('".addslashes(get_option('zopimBubbleText'))."');");
	}

	if (get_option('zopimBubbleEnable') == "show")
		echo "\n\$zopim.livechat.bubble.show(true);";
	else if (get_option('zopimBubbleEnable') == "hide")
		echo "\n\$zopim.livechat.bubble.hide(true);";

	// this must be called last
  if (get_option('zopimHideOnOffline') == "zopimHideOnOffline")
	 echo "\n\$zopim.livechat.button.setHideWhenOffline(true);";
  echo "\n})</script>";
}

function zopim_create_menu() {
  //create new top-level menu
  add_menu_page('Account Configuration', 'Zopim Chat', 'administrator', 'zopim_account_config', 'zopim_account_config', ZOPIM_SMALL_LOGO);

  // add_submenu_page('zopim_about', "About", "About", "administrator", 'zopim_about', 'zopim_about');
  add_submenu_page('zopim_account_config', 'Account Configuration', 'Account Setup', 'administrator', 'zopim_account_config', 'zopim_account_config');
  add_submenu_page('zopim_account_config', 'Customize Widget', 'Customize', 'administrator', 'zopim_customize_widget', 'zopim_customize_widget');
  add_submenu_page('zopim_account_config', 'IM Integration', 'IM Chat Bots', 'administrator', 'zopim_instant_messaging', 'zopim_instant_messaging');
  add_submenu_page('zopim_account_config', 'Dashboard', 'Dashboard', 'administrator', 'zopim_dashboard', 'zopim_dashboard');

  //call register settings function
  add_action( 'admin_init', 'register_mysettings' );
}

function check_zopimCode() {
/*
	//   if (get_option('zopimCode') == '' && ($_GET["page"] != "zopim_account_config")) {
	if (ereg("zopim", $_GET["page"] )) {
		//add_action( 'admin_notices', create_function( '', 'echo "<div class=\"error\"><p>" . sprintf( "Please <a href=\"%s\">input your Zopim account details</a>.", "admin.php?page=zopim_account_config" ) . "</p></div>";' ) );
		add_action( 'admin_notices', create_function( '', 'echo "<div class=\"error\"><p>This Zopim plugin is a work in progress. We will launch on the 25th of January. Thank you for your interest.</p></div>";' ) );
	}
 */
  return false;
}

function zopim_loader() {
  add_action( 'admin_menu', 'check_zopimCode' );
}

add_action( 'init', 'zopim_loader' );

function zopim_about() {
  echo "about";
}

function zopim_dashboard() {
  echo '<div id="dashboarddiv"><iframe id="dashboardiframe" src="'.ZOPIM_DASHBOARD_URL.'" height=700 width=98% scrolling="no"></iframe></div>      You may also <a href="'.ZOPIM_DASHBOARD_URL.'" target="_newWindow" onClick="javascript:document.getElementById(\'dashboarddiv\').innerHTML=\'\'; ">access the dashboard in a new window</a>.
		';
}

// Register the option settings we will be using
function register_mysettings() {

  // Authentication and codes
  register_setting( 'zopim-settings-group', 'zopimCode' );
  register_setting( 'zopim-settings-group', 'zopimUsername' );
  register_setting( 'zopim-settings-group', 'zopimSalt' );
  register_setting( 'zopim-settings-group', 'zopimUseSSL' );
  // General Widget settings
  register_setting( 'zopim-settings-group', 'zopimGetVisitorInfo' );
  register_setting( 'zopim-settings-group', 'zopimLang' );
  // Chat button settings
  register_setting( 'zopim-settings-group', 'zopimPosition' );
  register_setting( 'zopim-settings-group', 'zopimHideOnOffline' );
  register_setting( 'zopim-settings-group', 'zopimBubbleTitle' );
  register_setting( 'zopim-settings-group', 'zopimBubbleText' );
  register_setting( 'zopim-settings-group', 'zopimBubbleEnable' );
  register_setting( 'zopim-settings-group', 'zopimUseBubble' );
  // Themes / Color
  register_setting( 'zopim-settings-group', 'zopimColor' );
  register_setting( 'zopim-settings-group', 'zopimTheme' );
  // Message Settings
  register_setting( 'zopim-settings-group', 'zopimGreetings' );
  register_setting( 'zopim-settings-group', 'zopimUseGreetings' );

  get_option('zopimCode')         == "" && update_option('zopimCode', "zopim");
  get_option('zopimBubbleTitle')  == "" && update_option('zopimBubbleTitle', "Questions?");
  get_option('zopimBubbleText')   == "" && update_option('zopimBubbleText', "Click here to chat with us!");
  get_option('zopimBubbleEnable') == "" && update_option('zopimBubbleEnable', "checked");
  get_option('zopimUseGreetings') == "" && update_option('zopimUseGreetings', "disabled");
  get_option('zopimUseBubble')    == "" && update_option('zopimUseBubble', "disabled");
  get_option('zopimGreetings')    == "" && update_option('zopimGreetings', '{"away":{"window":"If you leave a question or comment, our agents will be notified and will try to attend to you shortly =)","bar":"Click here to chat"},"offline":{"window":"We are offline, but if you leave your message and contact details, we will try to get back to you =)","bar":"Leave a message"},"online":{"window":"Leave a question or comment and our agents will try to attend to you shortly =)","bar":"Click here to chat"}}');
}

add_action('get_footer', 'zopimme');
// create custom plugin settings menu
add_action('admin_menu', 'zopim_create_menu');

function do_post_request($url, $_data, $optional_headers = null)
{
  if (get_option('zopimUseSSL') != "zopimUseSSL")
	 $url = str_replace("https", "http", $url);

  $data = array();

  while(list($n,$v) = each($_data)){
	 $data[] = urlencode($n)."=".urlencode($v);
  }

  $data = implode('&', $data);

  $ch = curl_init();

  curl_setopt($ch, CURLOPT_URL, $url);
  curl_setopt($ch, CURLOPT_POST, true);
  curl_setopt($ch, CURLOPT_POSTFIELDS, $data);
  curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);

  $response = curl_exec($ch);
  curl_close($ch);

  return $response;
}

function json_to_array($json) {
  require_once('JSON.php');
  $jsonparser = new Services_JSON();
  return ($jsonparser->decode($json));
}

function to_json($variable) {
  require_once('JSON.php');
  $jsonparser = new Services_JSON();
  return ($jsonparser->encode($variable));
}

function getAccountDetails($salt) {
  $salty = array("salt" => get_option('zopimSalt'));
  return json_to_array(do_post_request(ZOPIM_GETACCOUNTDETAILS_URL, $salty));
}

function curl_get_url($filename) {
  $ch = curl_init();
  $timeout = 5; // set to zero for no timeout
  curl_setopt ($ch, CURLOPT_URL, $filename);
  curl_setopt ($ch, CURLOPT_RETURNTRANSFER, 1);
  curl_setopt ($ch, CURLOPT_CONNECTTIMEOUT, $timeout);
  $file_contents = curl_exec($ch);
  curl_close($ch);
  return $file_contents;
}

?>
