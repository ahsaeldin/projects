<?php
/**
 * The template for displaying all pages.
 *
 * This is the template that displays all pages by default.
 * Please note that this is the WordPress construct of pages
 * and that other 'pages' on your WordPress site will use a
 * different template.
 *
 * @package WordPress
 * @subpackage Theme_1144
 * @since Theme 1144 1.0
 */

get_header(); ?>

      <div class="background-type-2">
         <div class="main-content">

            <div id="container">
               <div id="content" role="main">
      
      <?php if ( have_posts() ) while ( have_posts() ) : the_post(); ?>
      
                  <div id="post-<?php the_ID(); ?>" <?php post_class(); ?>>
                     <?php if ( is_front_page() ) { ?>
                        <h2 class="entry-title"><?php the_title(); ?></h2>
                     <?php } else { ?>
                        <h1 class="entry-title"><?php the_title(); ?></h1>
                     <?php } ?>
      
                     <div class="entry-content">
                        <?php the_content(); ?>
                        <?php wp_link_pages( array( 'before' => '<div class="page-link">' . __( 'Pages:', 'theme1144' ), 'after' => '</div>' ) ); ?>
                        <?php edit_post_link( __( 'Edit', 'theme1144' ), '<span class="edit-link">', '</span>' ); ?>
                     </div><!-- .entry-content -->
                  </div><!-- #post-## -->
      
                  <?php comments_template( '', true ); ?>
      
      <?php endwhile; ?>
      
               </div><!-- #content -->
            </div><!-- #container -->
      
         </div>
      </div>

<?php get_footer(); ?>
