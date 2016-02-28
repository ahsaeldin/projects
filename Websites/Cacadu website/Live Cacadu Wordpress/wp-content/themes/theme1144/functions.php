<?php
/**
 * template_name functions and definitions
 *
 * Sets up the theme and provides some helper functions. Some helper functions
 * are used in the theme as custom template tags. Others are attached to action and
 * filter hooks in WordPress to change core functionality.
 *
 * The first function, template_name_setup(), sets up the theme by registering support
 * for various features in WordPress, such as post thumbnails, navigation menus, and the like.
 *
 * When using a child theme (see http://codex.wordpress.org/Theme_Development and
 * http://codex.wordpress.org/Child_Themes), you can override certain functions
 * (those wrapped in a function_exists() call) by defining them first in your child theme's
 * functions.php file. The child theme's functions.php file is included before the parent
 * theme's file, so the child theme functions would be used.
 *
 * Functions that are not pluggable (not wrapped in function_exists()) are instead attached
 * to a filter or action hook. The hook can be removed by using remove_action() or
 * remove_filter() and you can attach your own function to the hook.
 *
 * We can remove the parent theme's hook only after it is attached, which means we need to
 * wait until setting up the child theme:
 *
 * <code>
 * add_action( 'after_setup_theme', 'my_child_theme_setup' );
 * function my_child_theme_setup() {
 *     // We are providing our own filter for excerpt_length (or using the unfiltered value)
 *     remove_filter( 'excerpt_length', 'template_name_excerpt_length' );
 *     ...
 * }
 * </code>
 *
 * For more information on hooks, actions, and filters, see http://codex.wordpress.org/Plugin_API.
 *
 * @package WordPress
 * @subpackage Theme_1144
 * @since Theme 1144 1.0
 */

/**
 * Set the content width based on the theme's design and stylesheet.
 *
 * Used to set the width of images and content. Should be equal to the width the theme
 * is designed for, generally via the style.css stylesheet.
 */
if ( ! isset( $content_width ) )
	$content_width = 640;

/** Tell WordPress to run template_name_setup() when the 'after_setup_theme' hook is run. */
add_action( 'after_setup_theme', 'template_name_setup' );

if ( ! function_exists( 'template_name_setup' ) ):
/**
 * Sets up theme defaults and registers support for various WordPress features.
 *
 * Note that this function is hooked into the after_setup_theme hook, which runs
 * before the init hook. The init hook is too late for some features, such as indicating
 * support post thumbnails.
 *
 * To override template_name_setup() in a child theme, add your own template_name_setup to your child theme's
 * functions.php file.
 *
 * @uses add_theme_support() To add support for post thumbnails and automatic feed links.
 * @uses register_nav_menus() To add support for navigation menus.
 * @uses add_custom_background() To add support for a custom background.
 * @uses add_editor_style() To style the visual editor.
 * @uses load_theme_textdomain() For translation/localization support.
 * @uses add_custom_image_header() To add support for a custom header.
 * @uses register_default_headers() To register the default custom header images provided with the theme.
 * @uses set_post_thumbnail_size() To set a custom post thumbnail size.
 *
 * @since Theme 1144 1.0
 */
function template_name_setup() {

	// This theme styles the visual editor with editor-style.css to match the theme style.
	add_editor_style();

	// This theme uses post thumbnails
	add_theme_support( 'post-thumbnails' );

	// Add default posts and comments RSS feed links to head
	add_theme_support( 'automatic-feed-links' );

	// Make theme available for translation
	// Translations can be filed in the /languages/ directory
	load_theme_textdomain( 'theme1144', TEMPLATEPATH . '/languages' );

	$locale = get_locale();
	$locale_file = TEMPLATEPATH . "/languages/$locale.php";
	if ( is_readable( $locale_file ) )
		require_once( $locale_file );

	// This theme uses wp_nav_menu() in one location.
	register_nav_menus( array(
		'primary' => __( 'Primary Navigation', 'theme1144' ),
	) );

	// This theme allows users to set a custom background
	add_custom_background();

	// Your changeable header business starts here
	define( 'HEADER_TEXTCOLOR', '' );
	// No CSS, just IMG call. The %s is a placeholder for the theme template directory URI.
	define( 'HEADER_IMAGE', '%s/images/headers/path.jpg' );

	// The height and width of your custom header. You can hook into the theme's own filters to change these values.
	// Add a filter to template_name_header_image_width and template_name_header_image_height to change these values.
	define( 'HEADER_IMAGE_WIDTH', apply_filters( 'template_name_header_image_width', 940 ) );
	define( 'HEADER_IMAGE_HEIGHT', apply_filters( 'template_name_header_image_height', 198 ) );

	// We'll be using post thumbnails for custom header images on posts and pages.
	// We want them to be 940 pixels wide by 198 pixels tall.
	// Larger images will be auto-cropped to fit, smaller ones will be ignored. See header.php.
	set_post_thumbnail_size( HEADER_IMAGE_WIDTH, HEADER_IMAGE_HEIGHT, true );

	// Don't support text inside the header image.
	define( 'NO_HEADER_TEXT', true );

	// Add a way for the custom header to be styled in the admin panel that controls
	// custom headers. See template_name_admin_header_style(), below.
	add_custom_image_header( '', 'template_name_admin_header_style' );

	// ... and thus ends the changeable header business.

	// Default custom headers packaged with the theme. %s is a placeholder for the theme template directory URI.
	register_default_headers( array(
		'berries' => array(
			'url' => '%s/images/headers/berries.jpg',
			'thumbnail_url' => '%s/images/headers/berries-thumbnail.jpg',
			/* translators: header image description */
			'description' => __( 'Berries', 'theme1144' )
		),
		'cherryblossom' => array(
			'url' => '%s/images/headers/cherryblossoms.jpg',
			'thumbnail_url' => '%s/images/headers/cherryblossoms-thumbnail.jpg',
			/* translators: header image description */
			'description' => __( 'Cherry Blossoms', 'theme1144' )
		),
		'concave' => array(
			'url' => '%s/images/headers/concave.jpg',
			'thumbnail_url' => '%s/images/headers/concave-thumbnail.jpg',
			/* translators: header image description */
			'description' => __( 'Concave', 'theme1144' )
		),
		'fern' => array(
			'url' => '%s/images/headers/fern.jpg',
			'thumbnail_url' => '%s/images/headers/fern-thumbnail.jpg',
			/* translators: header image description */
			'description' => __( 'Fern', 'theme1144' )
		),
		'forestfloor' => array(
			'url' => '%s/images/headers/forestfloor.jpg',
			'thumbnail_url' => '%s/images/headers/forestfloor-thumbnail.jpg',
			/* translators: header image description */
			'description' => __( 'Forest Floor', 'theme1144' )
		),
		'inkwell' => array(
			'url' => '%s/images/headers/inkwell.jpg',
			'thumbnail_url' => '%s/images/headers/inkwell-thumbnail.jpg',
			/* translators: header image description */
			'description' => __( 'Inkwell', 'theme1144' )
		),
		'path' => array(
			'url' => '%s/images/headers/path.jpg',
			'thumbnail_url' => '%s/images/headers/path-thumbnail.jpg',
			/* translators: header image description */
			'description' => __( 'Path', 'theme1144' )
		),
		'sunset' => array(
			'url' => '%s/images/headers/sunset.jpg',
			'thumbnail_url' => '%s/images/headers/sunset-thumbnail.jpg',
			/* translators: header image description */
			'description' => __( 'Sunset', 'theme1144' )
		)
	) );
}
endif;

if ( ! function_exists( 'template_name_admin_header_style' ) ) :
/**
 * Styles the header image displayed on the Appearance > Header admin panel.
 *
 * Referenced via add_custom_image_header() in template_name_setup().
 *
 * @since Theme 1144 1.0
 */
function template_name_admin_header_style() {
?>
<style type="text/css">
/* Shows the same border as on front end */
#headimg {
	border-bottom: 1px solid #000;
	border-top: 4px solid #000;
}
/* If NO_HEADER_TEXT is false, you would style the text with these selectors:
	#headimg #name { }
	#headimg #desc { }
*/
</style>
<?php
}
endif;

/**
 * Get our wp_nav_menu() fallback, wp_page_menu(), to show a home link.
 *
 * To override this in a child theme, remove the filter and optionally add
 * your own function tied to the wp_page_menu_args filter hook.
 *
 * @since Theme 1144 1.0
 */
function template_name_page_menu_args( $args ) {
	$args['show_home'] = true;
	return $args;
}
add_filter( 'wp_page_menu_args', 'template_name_page_menu_args' );

/**
 * Sets the post excerpt length to 40 characters.
 *
 * To override this length in a child theme, remove the filter and add your own
 * function tied to the excerpt_length filter hook.
 *
 * @since Theme 1144 1.0
 * @return int
 */
function template_name_excerpt_length( $length ) {
	return 40;
}
add_filter( 'excerpt_length', 'template_name_excerpt_length' );

/**
 * Returns a "Continue Reading" link for excerpts
 *
 * @since Theme 1144 1.0
 * @return string "Continue Reading" link
 */
function template_name_continue_reading_link() {
	return ' <a href="'. get_permalink() . '" class="link-more">' . __( 'more', 'theme1144' ) . '</a>';
}

/**
 * Replaces "[...]" (appended to automatically generated excerpts) with an ellipsis and template_name_continue_reading_link().
 *
 * To override this in a child theme, remove the filter and add your own
 * function tied to the excerpt_more filter hook.
 *
 * @since Theme 1144 1.0
 * @return string An ellipsis
 */
function template_name_auto_excerpt_more( $more ) {
	return ' &hellip;' . template_name_continue_reading_link();
}
add_filter( 'excerpt_more', 'template_name_auto_excerpt_more' );

/**
 * Adds a pretty "Continue Reading" link to custom post excerpts.
 *
 * To override this link in a child theme, remove the filter and add your own
 * function tied to the get_the_excerpt filter hook.
 *
 * @since Theme 1144 1.0
 * @return string Excerpt with a pretty "Continue Reading" link
 */
function template_name_custom_excerpt_more( $output ) {
	if ( has_excerpt() && ! is_attachment() ) {
		$output .= template_name_continue_reading_link();
	}
	return $output;
}
add_filter( 'get_the_excerpt', 'template_name_custom_excerpt_more' );

/**
 * Remove inline styles printed when the gallery shortcode is used.
 *
 * Galleries are styled by the theme in Theme 1144's style.css.
 *
 * @since Theme 1144 1.0
 * @return string The gallery style filter, with the styles themselves removed.
 */
function template_name_remove_gallery_css( $css ) {
	return preg_replace( "#<style type='text/css'>(.*?)</style>#s", '', $css );
}
add_filter( 'gallery_style', 'template_name_remove_gallery_css' );

if ( ! function_exists( 'template_name_comment' ) ) :
/**
 * Template for comments and pingbacks.
 *
 * To override this walker in a child theme without modifying the comments template
 * simply create your own template_name_comment(), and that function will be used instead.
 *
 * Used as a callback by wp_list_comments() for displaying the comments.
 *
 * @since Theme 1144 1.0
 */
function template_name_comment( $comment, $args, $depth ) {
	$GLOBALS['comment'] = $comment;
	switch ( $comment->comment_type ) :
		case '' :
	?>
	<li <?php comment_class(); ?> id="li-comment-<?php comment_ID(); ?>">
		<div id="comment-<?php comment_ID(); ?>">
		<div class="comment-author vcard">
			<?php echo get_avatar( $comment, 40 ); ?>
			<?php printf( __( '%s <span class="says">says:</span>', 'theme1144' ), sprintf( '<cite class="fn">%s</cite>', get_comment_author_link() ) ); ?>
		</div><!-- .comment-author .vcard -->
		<?php if ( $comment->comment_approved == '0' ) : ?>
			<em><?php _e( 'Your comment is awaiting moderation.', 'theme1144' ); ?></em>
			<br />
		<?php endif; ?>

		<div class="comment-meta commentmetadata"><a href="<?php echo esc_url( get_comment_link( $comment->comment_ID ) ); ?>">
			<?php
				/* translators: 1: date, 2: time */
				printf( __( '%1$s at %2$s', 'theme1144' ), get_comment_date(),  get_comment_time() ); ?></a><?php edit_comment_link( __( '(Edit)', 'theme1144' ), ' ' );
			?>
		</div><!-- .comment-meta .commentmetadata -->

		<div class="comment-body"><?php comment_text(); ?></div>

		<div class="reply">
			<?php comment_reply_link( array_merge( $args, array( 'depth' => $depth, 'max_depth' => $args['max_depth'] ) ) ); ?>
		</div><!-- .reply -->
	</div><!-- #comment-##  -->

	<?php
			break;
		case 'pingback'  :
		case 'trackback' :
	?>
	<li class="post pingback">
		<p><?php _e( 'Pingback:', 'theme1144' ); ?> <?php comment_author_link(); ?><?php edit_comment_link( __('(Edit)', 'theme1144'), ' ' ); ?></p>
	<?php
			break;
	endswitch;
}
endif;

/**
 * Register widgetized areas, including two sidebars and four widget-ready columns in the footer.
 *
 * To override template_name_widgets_init() in a child theme, remove the action hook and add your own
 * function tied to the init hook.
 *
 * @since Theme 1144 1.0
 * @uses register_sidebar
 */
function template_name_widgets_init() {
	
	// Area 1, located in header.
	register_sidebar( array(
		'name' => __( 'Header Top Left Widget Area', 'theme1144' ),
		'id' => 'header-top-left-widget-area',
		'description' => __( 'The header widget area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
	// Area 2, located in header.
	register_sidebar( array(
		'name' => __( 'Header Top Right Widget Area', 'theme1144' ),
		'id' => 'header-top-right-widget-area',
		'description' => __( 'The header widget area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
	// Area 3, located at the top of the sidebar.
	register_sidebar( array(
		'name' => __( 'Primary Widget Area', 'theme1144' ),
		'id' => 'primary-widget-area',
		'description' => __( 'The primary widget area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
	// Area 4, located in contacts.php
	register_sidebar( array(
		'name' => __( 'Contacts Widget Area', 'theme1144' ),
		'id' => 'contacts-widget-area',
		'description' => __( 'Contacts Widget Area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
		// Area 5, located in faqs.php
	register_sidebar( array(
		'name' => __( 'FAQs Widget Area', 'theme1144' ),
		'id' => 'faqs-widget-area',
		'description' => __( 'Faqs Widget Area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
	// Area 6, located in support.php
	register_sidebar( array(
		'name' => __( 'Newsletter Widget Area', 'theme1144' ),
		'id' => 'newsletter-widget-area',
		'description' => __( 'Newsletter Widget Area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
	// Area 7, located in pricing.php
	register_sidebar( array(
		'name' => __( 'Pricing Widget Area', 'theme1144' ),
		'id' => 'pricing-widget-area',
		'description' => __( 'Pricing Widget Area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
	// Area 8, located in onlinehelp.php
	register_sidebar( array(
		'name' => __( 'Online Help Widget Area', 'theme1144' ),
		'id' => 'onlinehelp-widget-area',
		'description' => __( 'Online Help Widget Area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
	// Area 9, located in features.php
	register_sidebar( array(
		'name' => __( 'Testimonials Widget Area', 'theme1144' ),
		'id' => 'testimonials-widget-area',
		'description' => __( 'Testimonials Widget Area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
	// Area 10, located in search.php
	register_sidebar( array(
		'name' => __( 'Search Widget Area', 'theme1144' ),
		'id' => 'search-widget-area',
		'description' => __( 'Search Widget Area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );

	// Area 11, located in the footer. Empty by default.
	register_sidebar( array(
		'name' => __( 'First Footer Widget Area', 'theme1144' ),
		'id' => 'first-footer-widget-area',
		'description' => __( 'The first footer widget area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );

	// Area 12, located in the footer. Empty by default.
	register_sidebar( array(
		'name' => __( 'Second Footer Widget Area', 'theme1144' ),
		'id' => 'second-footer-widget-area',
		'description' => __( 'The second footer widget area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );

	// Area 13, located in the footer. Empty by default.
	register_sidebar( array(
		'name' => __( 'Third Footer Widget Area', 'theme1144' ),
		'id' => 'third-footer-widget-area',
		'description' => __( 'The third footer widget area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );

	// Area 14, located in the footer. Empty by default.
	register_sidebar( array(
		'name' => __( 'Fourth Footer Widget Area', 'theme1144' ),
		'id' => 'fourth-footer-widget-area',
		'description' => __( 'The fourth footer widget area', 'theme1144' ),
		'before_widget' => '<li id="%1$s" class="widget-container %2$s">',
		'after_widget' => '</li>',
		'before_title' => '<h3 class="widget-title">',
		'after_title' => '</h3>',
	) );
	
}
/** Register sidebars by running template_name_widgets_init() on the widgets_init hook. */
add_action( 'widgets_init', 'template_name_widgets_init' );

/**
 * Removes the default styles that are packaged with the Recent Comments widget.
 *
 * To override this in a child theme, remove the filter and optionally add your own
 * function tied to the widgets_init action hook.
 *
 * @since Theme 1144 1.0
 */
 
 // custom Post Types
add_action('init', 'custom_post_types');
function custom_post_types() {

	// Faded slider
		$labels = array(
			'name' => _x('Faded slider', 'post type general name'),
			'singular_name' => _x('Faded', 'post type singular name'),
			'add_new' => _x('Add Faded Slider', 'Faded'),
			'add_new_item' => __('Add New Slider'),
			'edit_item' => __('Edit Faded'),
			'new_item' => __('New Faded'),
			'view_item' => __('View Faded'),
			'search_items' => __('Search Faded'),
			'not_found' =>  __('No Custom found'),
			'not_found_in_trash' => __('No Faded found in Trash'),
			'parent_item_colon' => ''
		);
		register_post_type(
			'faded',
			array(
				'labels' => $labels,
				'public' => true,
				'show_ui' => true,
				'hierarchical' => false,
				
				'rewrite' => true,
				'exclude_from_search' => true,
				'show_in_nav_menus' => false,
				'supports' => array(
					'title',
					'excerpt',
					'editor',
					'thumbnail',
					'page-attributes',
					'comments',
				),
			)
		);
		
	// Info block
		$labels = array(
			'name' => _x('Custom posts', 'post type general name'),
			'singular_name' => _x('Custom posts', 'post type singular name'),
			'add_new' => _x('Add Custom post', 'Custom posts'),
			'add_new_item' => __('Add New Slider'),
			'edit_item' => __('Edit Custom posts'),
			'new_item' => __('New Custom posts'),
			'view_item' => __('View Custom posts'),
			'search_items' => __('Search Custom posts'),
			'not_found' =>  __('No Custom found'),
			'not_found_in_trash' => __('No Custom posts found in Trash'),
			'parent_item_colon' => ''
		);
		register_post_type(
			'custom-post',
			array(
				'labels' => $labels,
				'public' => true,
				'show_ui' => true,
				'hierarchical' => false,
				
				'rewrite' => true,
				'exclude_from_search' => true,
				'show_in_nav_menus' => false,
				'supports' => array(
					'title',
					'excerpt',
					'editor',
					'thumbnail',
					'page-attributes',
					'comments',
				),
			)
		);
		
	// Support block
		$labels = array(
			'name' => _x('Support posts', 'post type general name'),
			'singular_name' => _x('Support posts', 'post type singular name'),
			'add_new' => _x('Add Support post', 'Support posts'),
			'add_new_item' => __('Add New Slider'),
			'edit_item' => __('Edit Support posts'),
			'new_item' => __('New Support posts'),
			'view_item' => __('View Support posts'),
			'search_items' => __('Search Support posts'),
			'not_found' =>  __('No Support found'),
			'not_found_in_trash' => __('No Support posts found in Trash'),
			'parent_item_colon' => ''
		);
		register_post_type(
			'support-post',
			array(
				'labels' => $labels,
				'public' => true,
				'show_ui' => true,
				'hierarchical' => false,
				
				'rewrite' => true,
				'exclude_from_search' => true,
				'show_in_nav_menus' => false,
				'supports' => array(
					'title',
					'excerpt',
					'editor',
					'thumbnail',
					'page-attributes',
					'comments',
				),
			)
		);
		
	// Help Center
		$labels = array(
			'name' => _x('Help post', 'post type general name'),
			'singular_name' => _x('Help post', 'post type singular name'),
			'add_new' => _x('Add Help post', 'Support posts'),
			'add_new_item' => __('Add New Slider'),
			'edit_item' => __('Edit Help post'),
			'new_item' => __('New Help post'),
			'view_item' => __('View Help post'),
			'search_items' => __('Search Help posts'),
			'not_found' =>  __('No Help found'),
			'not_found_in_trash' => __('No Help posts found in Trash'),
			'parent_item_colon' => ''
		);
		register_post_type(
			'help-center',
			array(
				'labels' => $labels,
				'public' => true,
				'show_ui' => true,
				'hierarchical' => false,
				
				'rewrite' => true,
				'exclude_from_search' => true,
				'show_in_nav_menus' => false,
				'supports' => array(
					'title',
					'excerpt',
					'editor',
					'thumbnail',
					'page-attributes',
					'comments',
				),
			)
		);
		
		// Feature block
		$labels = array(
			'name' => _x('Feature posts', 'post type general name'),
			'singular_name' => _x('Feature posts', 'post type singular name'),
			'add_new' => _x('Add Feature post', 'Feature posts'),
			'add_new_item' => __('Add New Slider'),
			'edit_item' => __('Edit Feature posts'),
			'new_item' => __('New Feature posts'),
			'view_item' => __('View Feature posts'),
			'search_items' => __('Search Feature posts'),
			'not_found' =>  __('No Feature found'),
			'not_found_in_trash' => __('No Feature posts found in Trash'),
			'parent_item_colon' => ''
		);
		register_post_type(
			'features',
			array(
				'labels' => $labels,
				'public' => true,
				'show_ui' => true,
				'hierarchical' => false,
				
				'rewrite' => true,
				'exclude_from_search' => true,
				'show_in_nav_menus' => false,
				'supports' => array(
					'title',
					'excerpt',
					'editor',
					'thumbnail',
					'page-attributes',
					'comments',
				),
			)
		);
		
		// Features Block
		$labels = array(
			'name' => _x('Feature block', 'post type general name'),
			'singular_name' => _x('Feature block', 'post type singular name'),
			'add_new' => _x('Add Feature block', 'Features posts'),
			'add_new_item' => __('Add New Slider'),
			'edit_item' => __('Edit Feature block'),
			'new_item' => __('New Feature block'),
			'view_item' => __('View Feature block'),
			'search_items' => __('Search Feature block'),
			'not_found' =>  __('No Feature block'),
			'not_found_in_trash' => __('No Feature block found in Trash'),
			'parent_item_colon' => ''
		);
		register_post_type(
			'features-block',
			array(
				'labels' => $labels,
				'public' => true,
				'show_ui' => true,
				'hierarchical' => false,
				
				'rewrite' => true,
				'exclude_from_search' => true,
				'show_in_nav_menus' => false,
				'supports' => array(
					'title',
					'excerpt',
					'editor',
					'thumbnail',
					'page-attributes',
					'comments',
				),
			)
		);
		
		// Testimonials
		$labels = array(
			'name' => _x('Testimonials', 'post type general name'),
			'singular_name' => _x('Testimonial posts', 'post type singular name'),
			'add_new' => _x('Add Testimonial post', 'Feature posts'),
			'add_new_item' => __('Add New Testimonial'),
			'edit_item' => __('Edit Testimonial posts'),
			'new_item' => __('New Testimonial posts'),
			'view_item' => __('View Testimonial posts'),
			'search_items' => __('Search Testimonial posts'),
			'not_found' =>  __('No Testimonial found'),
			'not_found_in_trash' => __('No Testimonial posts found in Trash'),
			'parent_item_colon' => ''
		);
		register_post_type(
			'testimonial',
			array(
				'labels' => $labels,
				'public' => true,
				'show_ui' => true,
				'hierarchical' => false,
				
				'rewrite' => true,
				'exclude_from_search' => true,
				'show_in_nav_menus' => false,
				'supports' => array(
					'title',
					'excerpt',
					'editor',
					'thumbnail',
					'page-attributes',
					'comments',
				),
			)
		);
		
}
 
function template_name_remove_recent_comments_style() {
	global $wp_widget_factory;
	remove_action( 'wp_head', array( $wp_widget_factory->widgets['WP_Widget_Recent_Comments'], 'recent_comments_style' ) );
}
add_action( 'widgets_init', 'template_name_remove_recent_comments_style' );

if ( ! function_exists( 'template_name_posted_on' ) ) :
/**
 * Prints HTML with meta information for the current post—date/time and author.
 *
 * @since Theme 1144 1.0
 */
function template_name_posted_on() {
	printf( __( '%2$s', 'theme1144' ),
		'meta-prep meta-prep-author',
		sprintf( '<a href="%1$s" title="%2$s" rel="bookmark"><span class="entry-date">%3$s</span></a>',
			get_permalink(),
			esc_attr( get_the_time() ),
			get_the_date()
		),
		sprintf( '<span class="author vcard"><a class="url fn n" href="%1$s" title="%2$s">%3$s</a></span>',
			get_author_posts_url( get_the_author_meta( 'ID' ) ),
			sprintf( esc_attr__( 'View all posts by %s', 'theme1144' ), get_the_author() ),
			get_the_author()
		)
	);
}
endif;

if ( ! function_exists( 'template_name_posted_in' ) ) :
/**
 * Prints HTML with meta information for the current post (category, tags and permalink).
 *
 * @since Theme 1144 1.0
 */
function template_name_posted_in() {
	// Retrieves tag list of current post, separated by commas.
	$tag_list = get_the_tag_list( '', ', ' );
	if ( $tag_list ) {
		$posted_in = __( 'This entry was posted in %1$s and tagged %2$s. Bookmark the <a href="%3$s" title="Permalink to %4$s" rel="bookmark">permalink</a>.', 'theme1144' );
	} elseif ( is_object_in_taxonomy( get_post_type(), 'category' ) ) {
		$posted_in = __( 'This entry was posted in %1$s. Bookmark the <a href="%3$s" title="Permalink to %4$s" rel="bookmark">permalink</a>.', 'theme1144' );
	} else {
		$posted_in = __( 'Bookmark the <a href="%3$s" title="Permalink to %4$s" rel="bookmark">permalink</a>.', 'theme1144' );
	}
	// Prints the string, replacing the placeholders.
	printf(
		$posted_in,
		get_the_category_list( ', ' ),
		$tag_list,
		get_permalink(),
		the_title_attribute( 'echo=0' )
	);
}
endif;
