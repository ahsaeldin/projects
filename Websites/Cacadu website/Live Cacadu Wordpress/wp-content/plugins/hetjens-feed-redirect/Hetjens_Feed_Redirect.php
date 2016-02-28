<?php
/*
Plugin Name: Hetjens Feed Redirect
Plugin URI: http://hetjens.com/wordpress/hetjens_feed_redirect/
Version: 0.4
Description: Forwards the main feeds to Feedburner or any similar service.
Author: Philip Hetjens
Author URI: http://hetjens.com
Text Domain: Hetjens_Feed_Redirect
License: GPL
*/

/*
  Copyright 2010 Philip Hetjens (email : Philip at Hetjens dot eu);

  This program is free software; you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation; either version 3 of the License.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program; if not, write to the Free Software
  Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
*/

class Hetjens_Feed_Redirect {

  /*
   * Registers this plugin
   */
  function register() {
    add_action('template_redirect', array(&$this,'redirect'));
    add_action('admin_init', array(&$this,'register_admin'));
  }
  
  /*
   * Registers the wp-admin functions
   */
  function register_admin() {
    load_plugin_textdomain('hetjens_feed_redirect', false, '/'.dirname(plugin_basename(__FILE__)));

    add_settings_section('hetjens_feed_redirect', __('Feedburner settings','hetjens_feed_redirect'), array(&$this,'admin_section_description'), 'reading');
  
    add_settings_field('hetjens_feed_redirect_feed_url', '<label for="hetjens_feed_redirect_feed_url">'.__('Feedburner\'s Feed-URL','hetjens_feed_redirect').'</label>', array(&$this,'admin_feed_url'), 'reading', 'hetjens_feed_redirect', array());
    register_setting('reading','hetjens_feed_redirect_feed_url');

    add_settings_field('hetjens_feed_redirect_comments_url', '<label for="hetjens_feed_redirect_comments_url">'.__('Comment-Feed-URL','hetjens_feed_redirect').'</label>', array(&$this,'admin_comments_url'), 'reading', 'hetjens_feed_redirect', array());
    register_setting('reading','hetjens_feed_redirect_comments_url');
  }
  
  /*
   * Functions used by the settings api
   */
  function admin_section_description() {}
  function admin_feed_url() {
    $url = get_option('hetjens_feed_redirect_feed_url');
    echo '<input type="text" class="regular-text" id="hetjens_feed_redirect_feed_url" name="hetjens_feed_redirect_feed_url" value="'.$url.'" />';
  }
  function admin_comments_url() {
    $curl = get_option('hetjens_feed_redirect_comments_url');
    echo '<input type="text" class="regular-text" id="hetjens_feed_redirect_comments_url" name="hetjens_feed_redirect_comments_url" value="'.$curl.'" />';
  }
  
  /*
   * (maybe) forwards the feed
   */ 
  function redirect() {
    global $wp_query;

    // General checking
    if ((is_feed() == true) && (isset($_GET['noredirect']) == false) && !preg_match("/feedburner|feedvalidator/i", $_SERVER['HTTP_USER_AGENT'])) {
            
      //Normal feed
      $url = get_option('hetjens_feed_redirect_feed_url');
      if (($wp_query->is_comment_feed == false) && (trim($url) != '') && (is_archive() == false)) {
        status_header(302);
        header('Location: '.$url);
        die;
      }
  
      //Comment feed
      $curl = get_option('hetjens_feed_redirect_comments_url');
      if (($wp_query->is_comment_feed == true) && (trim($curl) != '') && (is_single() == false)) {
        status_header(302);
        header('Location: '.$curl);
        die;
      }
    }
  }
  
}

/* Initialise outselves */
add_action('plugins_loaded', create_function('','$Hetjens_Feed_Redirect = new Hetjens_Feed_Redirect(); $Hetjens_Feed_Redirect->register();'));
?>