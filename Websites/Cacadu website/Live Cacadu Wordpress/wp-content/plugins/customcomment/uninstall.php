<?php 
/**
 * @author Moein 
 * @copyright 2011
 *
 * The uninstallation script.
 */

if( defined( 'ABSPATH') && defined('WP_UNINSTALL_PLUGIN') ) {
	$CComment_fields = explode(",", get_option('CComment_fields'));
	//Unset the options when the plugin is deleted
	delete_option('CComment_fields');
	
	$all_comments = explode(",", get_option('CComment_comments_id'));
	//Remove all the post_meta that the plugin has been created
	foreach ($CComment_fields as $CComment_field){
		foreach( $all_comments as $comment_id)
				delete_post_meta($comment_id, "CuCo_".$CComment_field);
		//Remove extra field's data
		delete_option($CComment_field);
	}
}
?>