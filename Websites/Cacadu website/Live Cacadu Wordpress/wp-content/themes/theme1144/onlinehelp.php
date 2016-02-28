<?php
/**
 * Template Name: Online Help
 *
 * A custom page template without sidebar.
 *
 * The "Template Name:" bit above allows this to be selectable
 * from a dropdown menu on the edit page screen.
 *
 * @package WordPress
 * @subpackage Theme_1144
 * @since Theme 1144 1.0
 */

get_header(); ?>

      <div class="background-type-1">
         <div class="main-content">
         	<div class="line-ver">
               <div class="wrapper">
               
               	<?php get_sidebar('onlinehelp'); ?>

                  <div id="container" class="one-column">
                     <div id="content" role="main">
            
            <?php if ( have_posts() ) while ( have_posts() ) : the_post(); ?>
            
                        <div id="post-<?php the_ID(); ?>" <?php post_class(); ?>>
                           <h1 class="entry-title"><?php the_title(); ?></h1>
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
         </div>
      </div>

<?php get_footer(); ?>
