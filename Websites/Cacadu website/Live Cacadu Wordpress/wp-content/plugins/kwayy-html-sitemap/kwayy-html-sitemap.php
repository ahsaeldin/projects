<?php
/*
Plugin Name: Kwayy HTML Sitemap
Plugin URI: http://www.kwayyinfotech.com/our-work/kwayy-html-sitemap/
Description: Kwayy HTML Sitemap will generate HTML (not XML) sitemap for your sitemap page. The plugin will not only show Page and Posts but also your other Custom Post Type like Products etc. top. You can also configure to show or hide your Post Types. You just need to create a page for Sitemap and insert our shortcode <code>[kwayy-sitemap]</code> to display HTML sitemap. You can get support at http://forum.kwayyinfotech.com/
Version: 2.0
Author: Kwayy Infotech
Author URI: http://www.kwayyinfotech.com/
License: GPL2
*/




/*** Redirect on activation ***/
register_activation_hook(__FILE__, 'kwayyhs_activate');
add_action('admin_init', 'kwayyhs_redirect');
function kwayyhs_activate(){
    add_option('kwayyhs_do_activation_redirect', true);
}
function kwayyhs_redirect(){
    if (get_option('kwayyhs_do_activation_redirect', false)){
        delete_option('kwayyhs_do_activation_redirect');
        
        kwayyhs_set_default_option();
        
        wp_redirect('options-general.php?page=kwayyhs');
    }
}
/*** ***/


if( isset($_POST['kwayyhs-update']) ){
	add_action( 'admin_notices', 'kwayyhs_theme_upgrade_notice' );
}
function kwayyhs_theme_upgrade_notice() { ?>
	<div id="message" class="updated fade">
		<p>Kwayy HTML Sitemap options saved successfully. </p>
	</div>

<?php
}


function kwayyhs_set_default_option(){
	$list       = array();
	$post_types = kwayyhs_post_types();
	
	foreach ($post_types  as $post_type ){
		$list[] = $post_type->name;
		add_option('kwayyhs_active_'.$post_type->name , 'active' );
	}
	$list = implode(',', $list);
	add_option('kwayyhs_sortorder' , $list ); // storing sort order
}






function kwayyhs_post_types(){
	// http://codex.wordpress.org/Function_Reference/get_post_types
	$args=array(
	  'public'   => true
	  //'_builtin' => false
	);
	
	$output     = 'objects'; // names or objects, note names is the default
	$operator   = 'and'; // 'and' or 'or'
	$post_types = get_post_types($args,$output,$operator); 
	
	// Removing Attachment Custom Post Type
	unset($post_types["attachment"]);
	
	return $post_types;
}




add_action( 'admin_init', 'kwayyhs_init', 1 );
add_action( 'admin_menu', 'kwayyhs_adminbar_menu' );
add_action( 'plugin_action_links_' . plugin_basename(__FILE__), 'kwayyhs_plugin_actions');



function kwayyhs_plugin_actions($links){
	$new_links = array();
	$adminlink = get_bloginfo('url').'/wp-admin/';
	$fcmlink = 'http://www.fischercreativemedia.com/wordpress-plugins';
	$new_links[] = '<a href="'.$adminlink.'options-general.php?page=kwayyhs">Settings</a>';
	return array_merge($links,$new_links );
}


function kwayyhs_adminbar_menu(){
	if(is_multisite() && is_super_admin()){
		add_options_page( 'Kwayy HTML Sitemap Options', 'Kwayy HTML Sitemap Options','manage_network', 'kwayyhs', 'kwayyhs_page' );
	}elseif(is_multisite() && !is_super_admin()){
	    $theRoles = get_option('global-admin-bar-roles');
	    if(!is_array($theRoles)){$theRoles = array();}
	    if(!in_array(get_current_user_role(),$theRoles)){
			add_options_page( 'Kwayy HTML Sitemap Options', 'Kwayy HTML Sitemap Options','manage_options', 'kwayyhs', 'kwayyhs_page' );
		}
	}elseif(!is_multisite() && current_user_can('manage_options')){
		add_options_page( 'Kwayy HTML Sitemap Options', 'Kwayy HTML Sitemap Options','manage_options', 'kwayyhs', 'kwayyhs_page' );
	}
}



function kwayyhs_page(){
	
	// storing plugin options as array
	if( isset($_POST['kwayyhs-update']) ){
		update_option( 'kwayyhs_sortorder' , $_POST['kwayyhs-sortorder'] );
		update_option( 'kwayyhs_exclude' , $_POST['kwayyhs-exclude'] );
		
		$post_types2  = kwayyhs_post_types();
		
		foreach ( $post_types2 as $post_type ){
			if( isset( $_POST['kwayyhs_active_'.$post_type->name] ) ){
				update_option('kwayyhs_active_'.$post_type->name, 'active' );
			} else {
				update_option('kwayyhs_active_'.$post_type->name, 'deactive' );
			}
		}
	}
	
	// Retrive all options
	$kwayyhs_sortorder = get_option('kwayyhs_sortorder');
	$kwayyhs_exclude   = get_option('kwayyhs_exclude');
	?>
	
	
	<div class="wrap">
		<div id="icon-options-general" class="icon32"><br></div> <h2>Kwayy HTML Sitemap Options</h2>
		<br />
	
		<form method="post" action="">
			<input type="hidden" name="kwayyhs-update" id="kwayyhs-update" value="y" />
		
			<?php
			settings_fields( 'kwayyhs' );
			kwayyhs_set_default_option();
			?>
    
			
			<div id="poststuff">
				<div id="post-body" class="metabox-holder columns-2" style="min-width:650px;">
					<div id="post-body-content">
						<div id="kwayyhs-option-wrapper">
							<div id="kwayyhs-option-title">
								<div class="kwayyhs-title kwayyhs-title1">Drag to Sort</div>
								<div class="kwayyhs-title kwayyhs-title2">Show</div>
								<div class="kwayyhs-title kwayyhs-title3">Custom Post Name</div>
								<div class="kwayyhs-title kwayyhs-title4">Custom Post SLUG</div>
							</div><!-- #kwayyhs-option-title -->
				
							<ul id="kwayyhs-sortable">
								<?php
								$allposttype     = array();
								$post_types      = kwayyhs_post_types();
								$post_list_array = kwayyhs_post_list();
								echo kwayyhs_sortableList( $post_list_array );
								
								// creating sort order adding new post tye and removing removed post type
								$kwayyhs_sortorder = implode( ',',array_keys($post_list_array) );
								
								?>
							</ul><!-- #sortable -->
					
					
					
					
						</div><!-- #kwayyhs-option-wrapper -->
					
						
						<br />
					
					
						<div class="postbox" style="width:650px;" >
							<h3 class='hndle'>Exclude Post</h3>
							<div class="inside">
								<div class="submitbox">
									Exclude post:
									<input type="text" name="kwayyhs-exclude" id="kwayyhs-exclude" style="width:400px;" value="<?php echo $kwayyhs_exclude; ?>" />
									<p class="description">Please insert comma separated page IDs which you want to hide on Sitemap page. <br> Example: <code>8,56,98,106</code></p>
									<div class="clear"></div>
								</div><!-- .submitbox -->
							</div><!-- .inside -->
						</div><!-- #postbox-container-1 .postbox-container -->
					
					
					
					
					
					
					
					
					
						
				
						<input type="hidden" name="kwayyhs-sortorder" id="kwayyhs-sortorder" value="<?php echo $kwayyhs_sortorder; ?>" />
						<p class="submit"><input type="submit" class="button-primary" value="<?php _e('Save Changes') ?>" /></p>
				
					</div><!-- #post-body-content -->
			
			
			
					<div id="postbox-container-1" class="postbox-container">
						<div class="postbox">
							<h3 class='hndle'>Quick Guide</h3>
							<div class="inside">
								<div class="submitbox">
									Steps:
									<ol>
										<li>Select Custom Post Type from left, which you want to show on Sitemap Page. Than click "Save Changes" button.</li>
										<li>Create a new page (for sitemap) and insert <code>[kwayy-sitemap]</code> in content area.</li>
										<li>Done, yeah that's easy.</li>
									</ol>
									<hr>
									<div style="padding:10px; text-align:center;">
									<a href="http://forum.kwayyinfotech.com/categories/kwayy-html-sitemap" target="_blank" class="button-primary" >Get Help</a> &nbsp; &nbsp; &nbsp; <a href="http://forum.kwayyinfotech.com/categories/kwayy-html-sitemap" target="_blank" class="button">Report Bug</a>
									</div>
									<div class="clear"></div>
								
								</div><!-- .submitbox -->
							</div><!-- .inside -->
						</div><!-- .postbox -->
					</div><!-- .postbox-container #postbox-container-1 -->
					
					
					
					<div class="clear"></div>
				</div><!-- #post-body -->
				<div class="clear"></div>
			
			</div><!-- #poststuff -->
			
			
			
			
		</form>
		
	</div><!-- .wrap -->
	
	<?php
}




function kwayyhs_sortableList($post_types){
	$return  = '';
	
	foreach($post_types as $post_type){
		$checked = '';
		if($post_type->kwayy_active == 'yes' ){
			$checked = ' checked="checked" ';
		}
		
		$return .= '
		<li class="kwayyhs-ui-state-default" id="' . $post_type->name . '">
			<div class="kwayyhs-cpt">
				<div class="kwayyhs-dragable-handler"></div>
				<div class="kwayyhs-dragable-checkbox"><input name="kwayyhs_active_'.$post_type->name.'" id="kwayyhs_active_'.$post_type->name.'" type="checkbox" ' . $checked . ' /></div>
				<div class="kwayyhs-cpt-name">' . $post_type->labels->name . '</div>
				<div class="kwayyhs-cpt-slug">' . $post_type->name . '</div>
				<div class="clr"></div>
			</div>
		</li>
		';
	}
	return $return;
}



function kwayyhs_init(){
	wp_enqueue_script('jquery-ui-core');
	wp_enqueue_script('jquery-ui-sortable');
	wp_enqueue_script('kwayyhs-custom-js' ,  plugins_url( 'kwayy-html-sitemap.js', __FILE__ ) );
	wp_enqueue_style('kwayyhs-custom-css' ,  plugins_url( 'kwayy-html-sitemap.css', __FILE__ ) );
}




function sortArrayByArray($array,$orderArray){
	$ordered = array();
	foreach($orderArray as $key) {
		if(array_key_exists($key,$array)){
			if( get_option('kwayyhs_active_'.$key) == 'active' ){
				$array[$key]->kwayy_active = 'yes';
			}
			$ordered[$key] = $array[$key];
			unset($array[$key]);
		}
	}
	return $ordered + $array;
}



function kwayyhs_post_list(){
	$allposttype       = array();
	$new_allposttype   = array();
	$post_types        = kwayyhs_post_types();
	$kwayyhs_sortorder = get_option('kwayyhs_sortorder');
	
	
	$kwayyhs_sortorder_array = explode( ',', $kwayyhs_sortorder );
	$allposttype             = sortArrayByArray($post_types, $kwayyhs_sortorder_array);

	return $allposttype;
}




/******************* SHORTCODE *********************/
//[kwayy-sitemap]
function shortcode_kwayy_sitemap( $atts ){
	$return            = '<div class="kwayy-html-sitemap-wrapper">';
	$post_types        = kwayyhs_post_types();
	$kwayyhs_sortorder = get_option('kwayyhs_sortorder');
	//$kwayyhs_exclude   = get_option('kwayyhs_exclude');
	
	$kwayyhs_sortorder_array = explode( ',', $kwayyhs_sortorder );
	foreach($kwayyhs_sortorder_array as $post_type){
		if( get_option('kwayyhs_active_'.$post_type) == 'active' ){
			$return .= kwayyhs_get_post_by_post_type( $post_type , $post_types[$post_type]->labels->name );
		}
	}
	
	$return .= '</div> <!-- .kwayy-html-sitemap-wrapper -->';
	
	return $return;
}

add_shortcode( 'kwayy-sitemap', 'shortcode_kwayy_sitemap' );


function kwayyhs_get_post_by_post_type( $postype , $title , $orderby = 'menu_order' , $order = 'ASC' ){
	global $post;
	$curr_page_id = '';
	
	if( isset($post->ID) ){
		$curr_page_id = $post->ID;
	}
	
	$return = '';
	$args = array( 'post_type' => $postype, 'posts_per_page' => -1, 'orderby' => $orderby, 'order' => $order );
	$loop = new WP_Query( $args );
	wp_reset_query(); // Restting WP_Query
	
	$posts    = $loop->posts;
	$return .= '<h2 class="kwayy-html-sitemap-post-title kwayy-'.$loop->query_vars['post_type'].'-title">'.$title.'</h2>';
	
	if(count($posts) > 0 ){
		//echo '<pre>';
		//var_dump($loop->query_vars);
		//echo '</pre>';
		$return   .= '<ul class="kwayy-html-sitemap-post-list kwayy-'.$loop->query_vars['post_type'].'-list">';
		$parent_id = 0; // We are first start by fetching parent pages
		$return   .= kwayyhs_get_subpost( $posts , $parent_id , $curr_page_id );
		$return   .= '</ul>';
	}
	
	
	
	return $return;
}

function kwayyhs_get_subpost( $posts , $parent_id , $curr_page_id ){
	$return = '';
	$posts2 = $posts;
	
	$kwayyhs_exclude = get_option('kwayyhs_exclude');
	$kwayyhs_exclude = explode(',',$kwayyhs_exclude);
	//var_dump($kwayyhs_exclude);
	
	
	if( $posts > 0 ){
		foreach($posts as $post){
			
			if($post->post_parent == $parent_id){
				if( $post->ID != $curr_page_id ){
					if( !in_array($post->ID, $kwayyhs_exclude) ){
						$return .= '<li><a href="'.get_permalink($post->ID).'">'.$post->post_title.'</a>';
						$return .= kwayyhs_get_subpost( $posts2, $post->ID , $curr_page_id  );
						$return .= '</li>';
					}
				}
			}
		}
		if($return != ''){
			$return = '<ul>'.$return.'</ul>';
		}
	}
	
	return $return;
	
}

/***************************************************/

