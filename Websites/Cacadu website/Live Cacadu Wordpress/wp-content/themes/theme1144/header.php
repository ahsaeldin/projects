<?php
/**
 * The Header for our theme.
 *
 * Displays all of the <head> section and everything up till <div id="main">
 *
 * @package WordPress
 * @subpackage Theme_1144
 * @since Theme 1144 1.0
 */
?><!DOCTYPE html>
<html <?php language_attributes(); ?>>
<head>
<meta charset="<?php bloginfo( 'charset' ); ?>" />
<title><?php
	/*
	 * Print the <title> tag based on what is being viewed.
	 */
	global $page, $paged;

	wp_title( '|', true, 'right' );

	// Add the blog name.
	bloginfo( 'name' );

	// Add the blog description for the home/front page.
	$site_description = get_bloginfo( 'description', 'display' );
	if ( $site_description && ( is_home() || is_front_page() ) )
		echo " | $site_description";

	// Add a page number if necessary:
	if ( $paged >= 2 || $page >= 2 )
		echo ' | ' . sprintf( __( 'Page %s', 'theme1144' ), max( $paged, $page ) );

	?></title>
<link rel="profile" href="http://gmpg.org/xfn/11" />
<link rel="stylesheet" type="text/css" media="all" href="<?php bloginfo( 'stylesheet_url' ); ?>" />
<link rel="stylesheet" type="text/css" href="<?php bloginfo('stylesheet_directory'); ?>/css/superfish.css" />

<style type="text/css">@import url("<?php bloginfo('stylesheet_directory'); ?>/style.css");</style>
<link rel="alternate stylesheet" type="text/css" href="<?php bloginfo('stylesheet_directory'); ?>/css/color1_styles.css" title="color1_styles" media="screen" />
<link rel="alternate stylesheet" type="text/css" href="<?php bloginfo('stylesheet_directory'); ?>/css/color2_styles.css" title="color2_styles" media="screen" />
<link rel="alternate stylesheet" type="text/css" href="<?php bloginfo('stylesheet_directory'); ?>/css/color3_styles.css" title="color3_styles" media="screen" />

<link rel="pingback" href="<?php bloginfo( 'pingback_url' ); ?>" />
<!--[if lt IE 7]>
   <script type="text/javascript" src="http://info.template-help.com/files/ie6_warning/ie6_script_other.js"></script>
<![endif]-->
<!--[if IE]>
   <style type="text/css">
     #header .menu > li, #header .menu li ul, #faded ul li b a, .entry-content a.more-link, .link-more, #commentform input[type="submit"], div.wpcf7 input[type="submit"], .module, .module input[type="submit"], .list-1 dd, .list-1 dd b a, .list-3 li, .list-6 li {
         behavior:url(<?php bloginfo('stylesheet_directory'); ?>/PIE.php)
      }
   </style>
<![endif]-->
<script type="text/javascript" src="<?php bloginfo('stylesheet_directory'); ?>/js/jquery-1.4.4.min.js"></script>
<script type="text/javascript" src="<?php bloginfo('stylesheet_directory'); ?>/js/superfish.js"></script>
<script type="text/javascript">
	var $j = jQuery.noConflict();
	$j(document).ready(function() {
		$j('#header .menu').superfish({ 
			delay:       1000,
			animation:   {opacity:'show',height:'show'},
			speed:       'fast',
			autoArrows:  false,
			dropShadows: false
		}); 
	});
</script>
<script type="text/javascript" src="<?php bloginfo('stylesheet_directory'); ?>/js/jquery.faded.js"></script>
<script type="text/javascript">
	var $j = jQuery.noConflict();
	$j(function(){
		$j("#faded").faded({
			speed: 900,
			autoplay: 10000,
			pagination: false,
			autorestart: true
		});
	});
</script>
<script type="text/javascript" src="<?php bloginfo('stylesheet_directory'); ?>/js/accordion.js"></script>
<script type="text/javascript">
$j().ready(function(){
	$j('.accordion').accordion({
		 active: '.active',
		 selectedClass: 'active',
		 header: "dt"
	})
});
</script>
<script src="<?php bloginfo('stylesheet_directory'); ?>/js/styleswitch.js" type="text/javascript"></script>
<script src="<?php bloginfo('stylesheet_directory'); ?>/js/jquery.tabSlideOut.v1.2.js" type="text/javascript"></script>
<!--
<script type="text/javascript">
    $j(function(){
        $j('.slide-out-div').tabSlideOut({
            tabHandle: '.handle',                     //class of the element that will become your tab
            pathToTabImage: '<?php bloginfo('stylesheet_directory'); ?>/images/change.png', //path to the image for the tab //Optionally can be set using css
            imageHeight: '1480px',                     //height of tab image           //Optionally can be set using css
            imageWidth: '37px',                       //width of tab image            //Optionally can be set using css
            tabLocation: 'left',                      //side of screen where tab lives, top, right, bottom, or left
            speed: 300,                               //speed of animation
            action: 'click',                          //options: 'click' or 'hover', action to trigger animation
            topPos: '87px',                          //position from the top/ use if tabLocation is left or right
            leftPos: '20px',                          //position from left/ use if tabLocation is bottom or top
            fixedPosition: true                      //options: true makes it stick(fixed position) on scroll
        });

    });
</script>
-->
<script type="text/javascript">
	var $j = jQuery.noConflict();
	$j(document).ready(function() {
		$j(".list-3 li:even").addClass("even");
	});
</script>
<script type="text/javascript">
	var $j = jQuery.noConflict();
	$j(document).ready(function() {
		$j(".list-6 li:even").addClass("even");
	});
</script>
<script type="text/javascript">
	var $j = jQuery.noConflict();
	$j(document).ready(function() {
		$j('input[type="text"]').addClass("idleField");
				$j('input[type="text"]').focus(function() {
				 $j(this).removeClass("idleField").addClass("focusField");
				 if (this.value == this.defaultValue){ 
				  this.value = '';
		 }
		 if(this.value != this.defaultValue){
			  this.select();
			 }
			});
			$j('input[type="text"]').blur(function() {
			 $j(this).removeClass("focusField").addClass("idleField");
				 if ($j.trim(this.value) == ''){
			  this.value = (this.defaultValue ? this.defaultValue : '');
		 }
			});
	});
</script>

<?php
	/* We add some JavaScript to pages with the comment form
	 * to support sites with threaded comments (when in use).
	 */
	if ( is_singular() && get_option( 'thread_comments' ) )
		wp_enqueue_script( 'comment-reply' );

	/* Always have wp_head() just before the closing </head>
	 * tag of your theme, or you will break many plugins, which
	 * generally use this hook to add elements to <head> such
	 * as styles, scripts, and meta tags.
	 */
	wp_head();
?>
</head>

<body <?php body_class(); ?>>
<div id="wrapper" class="hfeed">
   
   <!--
   <div class="slide-out-div">
      <a class="handle" href="http://link-for-non-js-users.html">Content</a>
      <ul>
         <li><a href="#" class="color1 styleswitch" rel="color1_styles">Style green</a></li>
         <li><a href="#" class="color2 styleswitch" rel="color2_styles">Style blue</a></li>
         <li><a href="#" class="color3 styleswitch" rel="color3_styles">Style orange</a></li>
      </ul>
   </div>
   -->
	
   <div id="header-tail">
   	<div id="slider-bg">
         <div id="header">
            <div id="masthead">

					<?php
                  // The header top-left-widget area
                  if ( is_active_sidebar( 'header-top-left-widget-area' ) ) : ?>
               
                     <div id="header-widget-left" class="widget-area" role="complementary">
                        <ul class="xoxo">
                           <?php dynamic_sidebar( 'header-top-left-widget-area' ); ?>
                        </ul>
                     </div>
               
               <?php endif; ?>
               
               <?php
                  // The header top-right-widget area
                  if ( is_active_sidebar( 'header-top-right-widget-area' ) ) : ?>
               
                     <div id="header-widget-right" class="widget-area" role="complementary">
                        <ul class="xoxo">
                           <?php dynamic_sidebar( 'header-top-right-widget-area' ); ?>
                        </ul>
                     </div>
               
               <?php endif; ?>
            
               <div id="branding" role="banner">
               
                  <?php $heading_tag = ( is_home() || is_front_page() ) ? 'h1' : 'div'; ?>
                  <<?php echo $heading_tag; ?> id="site-title">
                     <span>
                      <a href="<?php echo home_url( '/' ); ?>" title="<?php echo esc_attr( get_bloginfo( 'name', 'display' ) ); ?>" rel="home"><img src="<?php bloginfo('stylesheet_directory'); ?>/images/logo.png" title="" alt="" /></a>
                     </span>
                  </<?php echo $heading_tag; ?>>
                  
               </div><!-- #branding -->
      
               <div id="access" class="#header .menu" role="navigation">
                 <?php /*  Allow screen readers / text browsers to skip the navigation menu and get right to the good stuff */ ?>
                  <div class="skip-link screen-reader-text"><a href="#content" title="<?php esc_attr_e( 'Skip to content', 'theme1144' ); ?>"><?php _e( 'Skip to content', 'theme1144' ); ?></a></div>
                  <?php /* Our navigation menu.  If one isn't filled out, wp_nav_menu falls back to wp_page_menu.  The menu assiged to the primary position is the one used.  If none is assigned, the menu with the lowest ID is used.  */ ?>
                  <?php wp_nav_menu( array( 'container_class' => 'menu-header', 'theme_location' => 'primary' ) ); ?>
               </div><!-- #access -->
               
               <div id="faded">
						<?php $loop = new WP_Query(array('post_type' => 'faded', 'posts_per_page' => 15)); ?>
                     <ul>
                        <?php if ($loop->have_posts()): ?>
                        <?php $posts_counter = 0; ?>
                        <?php while ($loop->have_posts()) : $loop->the_post(); $posts_counter++; ?>
                        <li id="slide-<?php echo $posts_counter; ?>">
                        	<div class="inner">
                              <span><a href="<?php the_permalink(); ?>"><?php the_post_thumbnail(''); ?></a></span>
                              <strong><a href="<?php the_permalink(); ?>"><?php the_title(''); ?></a></strong>
			  	              <?php the_excerpt(); ?>
<b><a href="https://conderella.com/downloads/Cacadu.exe"><img src="<?php bloginfo('stylesheet_directory'); ?>/images/download.png" title="" alt="" /></a></b>

                           </div>
                        </li>
                        <?php endwhile; ?>
                     </ul>
                     <!--Disable arrows
					 <a href="#" class="prev"></a>
                     <a href="#" class="next"></a>
					 -->
                  <?php endif; ?>
                  <?php wp_reset_query(); ?>
               </div>
               
               <div class="breadcrumbs">
						<?php if ( function_exists('yoast_breadcrumb') ) {
                     yoast_breadcrumb('<p id="breadcrumbs">','</p>');
                  } ?>
               </div>
               
            </div><!-- #masthead -->
         </div><!-- #header -->
      </div><!-- #slider-bg -->
   </div><!-- #header-tail -->
   

	<div id="main">