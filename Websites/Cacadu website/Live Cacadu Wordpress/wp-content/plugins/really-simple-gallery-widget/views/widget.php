<?php

echo $before_widget;

if ( !empty( $title ) ) { echo $before_title . $title . $after_title; }

// initialize the images array
$images = array();

// on single post and option for images from single post on
if ( $current_post_images == '1' && is_single() ) {
	global $post;
	$post_id = $post->ID;
}

// using full media library
elseif ( $full_library == '1' ) {
	// unset the post_id for the next if statement
	unset($post_id);

	$query_images_args = array(
		'post_type' => 'attachment',
		'post_mime_type' =>'image',
		'post_status' => 'inherit',
		'posts_per_page' => -1,
		'order' => $order,
	);
	
	// randomly ordered - no more shuffle_assoc!
	if ($order == 'rand') {
		$query_images_args['orderby'] = 'rand';
	}
	
	// limited number of images
	if ($num_images != 0) {
		$query_images_args['posts_per_page'] = $num_images;
	}
	
	$query_images = new WP_Query( $query_images_args );
	
	foreach ( $query_images->posts as $attachment) {
	  $image = wp_get_attachment_image_src($attachment->ID, $image_size);
	  $image['ID'] = $attachment->ID;
	  $image['title'] = $attachment->post_title;
	  $image['parent'] = $attachment->post_parent;
	  $images[] = $image;
	}
}

// using specific post ID or on single post/page
if ($post_id) {
	$query_images_args = array(
		'post_parent' => $post_id,
		'post_status' => 'inherit',
		'post_type' => 'attachment',
		'post_mime_type' => 'image',
		'numberposts' => -1,
		'order' => $order,
	);
	
	// randomly ordered
	if ($order == 'rand') {
		$query_images_args['orderby'] = 'rand';
	}
	
	// limited number of images
	if ($num_images != 0) {
		$query_images_args['numberposts'] = $num_images;
	}

  $attachments = get_children( $query_images_args );
  
  if ( !empty($attachments) ) {
  	foreach( $attachments as $id => $attachment ) {
  		$image = wp_get_attachment_image_src($id, $image_size);
  		$image['ID'] = $id;
  		$image['title'] = $attachment->post_title;
  		$images[] = $image;
		}
	}
}

// there's stuff in the images array, hooray! let's output them images
if ( !empty($images) ) {

	$output = "<div class='rsgallery'>";

	foreach( $images as $image ) {
		if ($before_link_title != '')
			$link_title = "$before_link_title $image[title]";
		else
			$link_title = $image['title'];
		
		$image_output = '<img src="' . $image[0] . '" alt="' . $link_title . '" title="' . $link_title . '" width="'.$image[1].'" height="'.$image[2].'" class="rsg_image" />';
		$link_output = '';
		
		$id = $image['ID'];
		
		// stupid int(0)
		if ( $image['parent'] || $image['parent'] === 0 )
			$post_id = $image['parent'];

		// no post parent - set this one to use a file link if links are being used
		// perhaps add an option for attachment later
		if ( $post_id == '0' ) {
			$temp_link_type = 'file';
		}
		
		if ( $link_type != 'none' ) {
		
			if ( $link_type == 'file' || $temp_link_type == 'file' )
				$link_url = wp_get_attachment_url($id);
			elseif ( $link_type == 'post' )
				$link_url = get_permalink($post_id);
			elseif ( $link_type == 'anchor' )
				$link_url = get_permalink($post_id)."#attachment_$id";
			elseif ($link_type == 'attachment')
				$link_url = get_attachment_link($id);
	
			$link_output .= "<a href='$link_url'";
			
			if ($link_rel != '')
				$link_output .= " rel='$link_rel'";
			
			$link_output .= " title='$link_title'>$image_output</a>";
		}
		else
			$link_output = $image_output;
		
		$output .= $link_output;
		
		if ($show_captions == '1') {
			$caption = $attachment->post_excerpt;
			$output .= "<span class='rsg_caption'>$caption</span>";
		}
		
		//$i++;
	} 
	
	$output .= "
	</div>\n";
	
	echo $output;
} // !empty($images)

echo $after_widget;
?>