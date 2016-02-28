<?php
/*
Plugin Name: Really Simple Gallery Widget
Plugin URI: http://www.helenhousandi.com/wordpress/plugins/really-simple-gallery-widget/
Description: Widget to display images attached to a specific page/post or from the entire media library. Includes options for number of images, display order (including true random), any registered image size, caption toggling, type of link, rel attribute, and link title prefix. Can also show images attached to the page/post currently being viewed.
Version: 1.1
Author: Helen Hou-Sandi
Author URI: http://www.helenhousandi.com
*/

// function for proper associative array shuffling
if ( !function_exists('shuffle_assoc') ) {
	function shuffle_assoc($list) {
	  if (!is_array($list)) return $list;
	
	  $keys = array_keys($list);
	  shuffle($keys);
	  $random = array();
	  foreach ($keys as $key)
	    $random[$key] = $list[$key];

	  return $random;
	}
}

if ( !class_exists('RSGWidget') ) {
	class RSGWidget extends WP_Widget {
	
	  function RSGWidget() {
	    $widget_ops = array('classname' => 'widget_rsg', 'description' => __('Grab photos from a specified post/page or the entire media library and display them in a widget area.'));
	    $this->WP_Widget('RSGWidget', __('Really Simple Gallery Widget'), $widget_ops);
	  }
		
	  function widget( $args, $instance ) {
	    extract($args);
	    $title = apply_filters('widget_title', empty($instance['title']) ? '' : esc_html($instance['title']), $instance);
	    $post_id = $instance['post_id'];
	    $full_library = $instance['full_library'] ? '1' : '0';
	    $current_post_images = $instance['current_post_images'] ? '1' : '0';
	    $num_images = $instance['num_images'];
	    $show_captions = $instance['show_captions'] ? '1' : '0';
	    $image_size = $instance['image_size'];
	    $order = $instance['order'];
	    $link_type = $instance['link_type'];
	    $before_link_title = esc_attr( strip_tags($instance['before_link_title']) );
	    $link_rel = esc_attr( strip_tags($instance['link_rel']) );
	    
	    include('views/widget.php');
	  }
	  
	  function update( $new_instance, $old_instance ) {
			$instance = $old_instance;
			$instance['title'] = esc_html( strip_tags($new_instance['title']) );
			if ( !ctype_digit($new_instance['post_id']) )
			  $instance['post_id'] =  '';
			else
			  $instance['post_id'] = intval($new_instance['post_id']);
			$instance['full_library'] = $new_instance['full_library'] ? 1 : 0;
			$instance['current_post_images'] = $new_instance['current_post_images'] ? 1 : 0;
		  if ( !ctype_digit($new_instance['num_images']) )
		    $instance['num_images'] =  '0';
		  else
		    $instance['num_images'] = intval($new_instance['num_images']);
			$instance['show_captions'] = $new_instance['show_captions'] ? 1 : 0;
			$instance['image_size'] = $new_instance['image_size'];
			$instance['order'] = $new_instance['order'];
			$instance['link_type'] = $new_instance['link_type'];
			$instance['before_link_title'] = esc_attr( strip_tags($new_instance['before_link_title']) );
			$instance['link_rel'] = esc_attr( strip_tags($new_instance['link_rel']) );
			return $instance;
	  }
	
	  function form( $instance ) {
			include('views/form.php');
		}
	}
	function RSGWidget_init() {
	    register_widget('RSGWidget');
	}
	
	add_action('widgets_init', 'RSGWidget_init');
}