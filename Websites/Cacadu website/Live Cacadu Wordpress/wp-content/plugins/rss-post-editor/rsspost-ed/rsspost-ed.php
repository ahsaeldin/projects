<?php
/*
Plugin Name: RSS Post Editor
Version: 0.1
Plugin URI: http://www.techytuts.com/uncategorized/wordpress-plugin-rss-post-editor.html
Description: Allows you to add content RSS feed articles so that users have to subscribe to your feeds to view such content. helping you increase your RSS Readers
Author: Mohd Ameenuddin Atif
Author URI: http://www.techytuts.com/
*/

/*
 * USAGE:
 *
 * While writing/editing your post
 * Wrap any html between [rss][/rss] tags to make it visible only on your feed (please note that [rss][/rss] are case sensitive)
 * Anything between [rss][/rss] will not be visible directly on your blog.
 * Example : This is the normal post content [rss]this content is to be shown inside the rss feed[/rss] 
 * 
 */

function check_rss($content) {
		if(is_feed()) {
			$content = preg_replace('/\[rss\](.*?)\[\/rss\]/','$1',$content);	
		} else {
			$content = preg_replace('/\[rss\](.*?)\[\/rss\]/','',$content);	
		}
	return $content;
}

add_filter('the_content', 'check_rss');

?>