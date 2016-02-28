<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title> Coming Soon </title>
<?php
header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
header("Cache-Control: no-store, no-cache, must-revalidate");
header("Cache-Control: post-check=0, pre-check=0", false);
header("Pragma: no-cache");
require_once( dirname(__FILE__) . '../../../../../../wp-load.php');
if ( !defined('WP_CONTENT_URL') )define( 'WP_CONTENT_URL', get_option('siteurl') . '/wp-content');
if ( !defined('WP_CONTENT_DIR') )define( 'WP_CONTENT_DIR', ABSPATH . 'wp-content' );
/***********************
 * PLUGIN DIRECTORY PATH
***********************/
$plugin_path = WP_CONTENT_URL.'/plugins/'.plugin_basename(dirname(__FILE__));
?>
<link rel="stylesheet" media="screen" href="<?php echo $plugin_path; ?>/style.css" />
</head>
<body>
<br /><br /><br />
<h2 align="center">Coming Soon</h2>
<br />
<p align="center">
    You can paste your xHTML/CSS/Javascript documents in <strong>wp-plugins/cj-coming-soon/themes/custom_xhtml/</strong> directory.
</p>
<p align="center">
    If you want to use the back-end options, copy all files from the default directory and modify as per your requirements.
</p>
</body>
</html>
<?php die(); ?>