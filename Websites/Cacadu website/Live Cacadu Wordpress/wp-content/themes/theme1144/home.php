<?php
/**
 * The main template file.
 *
 * This is the most generic template file in a WordPress theme
 * and one of the two required files for a theme (the other being style.css).
 * It is used to display a page when nothing more specific matches a query. 
 * E.g., it puts together the home page when no home.php file exists.
 * Learn more: http://codex.wordpress.org/Template_Hierarchy
 *
 * @package WordPress
 * @subpackage Theme_1144
 * @since Theme 1144 1.0
 */

get_header(); ?>
		
      <div class="background-type-1">
         <div class="custom-posts">

				<?php $loop = new WP_Query(array('post_type' => 'custom-post', 'posts_per_page' => 3)); ?>
				<?php if ($loop->have_posts()): ?>
                  <ul>
							<?php $posts_counter = 0; ?>
								<?php while ($loop->have_posts()) :	$loop->the_post(); $posts_counter++; ?>
                           <li class="post-<?php echo $posts_counter; ?>">
                              <h3 class="widget-title"><?php the_title(''); ?></h3>
                              <?php the_excerpt(''); ?>
                              <b><a href="<?php the_permalink(); ?>">Read More</a></b>
                           </li>
                        <?php endwhile; ?>
                  </ul>
            <?php endif; ?>

         </div>
      </div>
      
      <div class="background-type-2">
         <div class="main-content">
         	<div class="line-ver">
               <div class="wrapper">
            
                  <?php get_sidebar(); ?>
            
                  <div id="container">
                     <div id="content" role="main">
            
                  <?php $wp_query = new WP_Query('post_type=post&posts_per_page=1&paged='.$paged ); ?>
                  <?php while ( $wp_query->have_posts() ) : $wp_query->the_post(); ?>
               
                        <div id="post-<?php the_ID(); ?>" <?php post_class(); ?>>
                           
                           <div class="wrapper">
                              <div class="entry-meta">
                                 <span class="date"><?php the_time('F jS, Y'); ?></span>
                                 <span class="comments-link"><?php comments_popup_link( __( '0', 'theme1144' ), __( '1', 'theme1144' ), __( '%', 'theme1144' ) ); ?></span>
                              </div>
                       			<h2 class="entry-title"><a href="<?php the_permalink(); ?>" title="<?php printf( esc_attr__( 'Permalink to %s', 'theme1144' ), the_title_attribute( 'echo=0' ) ); ?>" rel="bookmark"><?php the_title(); ?></a></h2>
                              <div class="clear"></div>
                           </div>
                           
                           <div class="entry-content">
                           	<?php the_excerpt( __( '<span>more</span>', 'theme1144' ) ); ?>
                           </div><!-- .entry-content -->
                           
                           
                        </div><!-- #post-## -->
                     
                   <?php endwhile; ?>
                     
                     </div><!-- #content -->
                  </div><!-- #container -->
               
               </div>
            </div>
         </div>
      </div>
      
<?php get_footer(); ?>
