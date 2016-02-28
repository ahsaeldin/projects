<?php
/**
 * The Template for displaying all single posts.
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
               
                  <div class="wrapper">
                     <div class="entry-meta">
                        <span class="date"><?php the_time('F jS, Y'); ?></span>
                     </div>
                     <h2 class="entry-title"><?php the_title(); ?></h2>
                     <div class="clear"></div>
                  </div>
   
                  <div class="entry-content">
                     <?php the_content(); ?>
                     <?php wp_link_pages( array( 'before' => '<div class="page-link">' . __( 'Pages:', 'theme1144' ), 'after' => '</div>' ) ); ?>
                  </div><!-- .entry-content -->
   
   <?php if ( get_the_author_meta( 'description' ) ) : // If a user has filled out their description, show a bio on their entries  ?>
                  <div id="entry-author-info">
                     <div id="author-avatar">
                        <?php echo get_avatar( get_the_author_meta( 'user_email' ), apply_filters( 'template_name_author_bio_avatar_size', 60 ) ); ?>
                     </div><!-- #author-avatar -->
                     <div id="author-description">
                        <h2><?php printf( esc_attr__( 'About %s', 'theme1144' ), get_the_author() ); ?></h2>
                        <?php the_author_meta( 'description' ); ?>
                        <div id="author-link">
                           <a href="<?php echo get_author_posts_url( get_the_author_meta( 'ID' ) ); ?>">
                              <?php printf( __( 'View all posts by %s <span class="meta-nav">&rarr;</span>', 'theme1144' ), get_the_author() ); ?>
                           </a>
                        </div><!-- #author-link	-->
                     </div><!-- #author-description -->
                  </div><!-- #entry-author-info -->
   <?php endif; ?>
   
                  <div class="entry-utility">
                     <?php template_name_posted_in(); ?>
                     <?php edit_post_link( __( 'Edit', 'theme1144' ), '<span class="edit-link">', '</span>' ); ?>
                  </div><!-- .entry-utility -->
               </div><!-- #post-## -->
   
               <div id="nav-below" class="navigation">
                  <div class="nav-previous"><?php previous_post_link( '%link', '<span class="meta-nav">' . _x( '&larr; Previous post', 'theme1144' ) . '</span>' ); ?></div>
                  <div class="nav-next"><?php next_post_link( '%link', '<span class="meta-nav">' . _x( 'Next post &rarr;', 'theme1144' ) . '</span>' ); ?></div>
               </div><!-- #nav-below -->
   
               <?php comments_template( '', true ); ?>
   
   <?php endwhile; // end of the loop. ?>
   
            </div><!-- #content -->
         </div><!-- #container -->
      </div>
   </div>

<?php get_footer(); ?>
