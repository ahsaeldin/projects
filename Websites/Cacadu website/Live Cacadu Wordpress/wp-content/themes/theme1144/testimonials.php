<?php
/**
 * Template Name: Testimonials
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

      <div class="background-type-2">
         <div class="main-content">
         	<div class="line-ver">
               <div class="wrapper">
            
                  <?php get_sidebar(testimonials); ?>

                  <div id="container">
                     <div id="content" role="main">
            				
                        <div class="wrapper"><h2 class="entry-title">Testimonials</h2></div>
                        
                      <?php $wp_query = new WP_Query('post_type=testimonial&posts_per_page=10&paged='.$paged ); ?>
                    <?php $posts_counter = 0; ?>
                     <ul class="list-6">
                           <?php while ($wp_query->have_posts()) :$wp_query->the_post(); $posts_counter++; ?>
                              <li class="feature-<?php echo $posts_counter; ?>">
                                 <div>
                                    <div>
                                       <strong><?php the_time('d D, Y'); ?> <a href="<?php the_permalink(); ?>"><?php the_title(''); ?></a></strong>
                                       <?php the_excerpt(''); ?>
                                    </div>
                                 </div>
                              </li>
                           <?php endwhile; ?>
                     </ul>
                     
                     
							
<?php /* Display navigation to next/previous pages when applicable */ ?>
<?php if (  $wp_query->max_num_pages > 1 ) : ?>
				<div id="nav-below" class="navigation">
					<div class="nav-previous"><?php next_posts_link( __( '<span class="meta-nav">&larr;</span> Older posts', 'theme1144' ) ); ?></div>
					<div class="nav-next"><?php previous_posts_link( __( 'Newer posts <span class="meta-nav">&rarr;</span>', 'theme1144' ) ); ?></div>
				</div><!-- #nav-below -->
<?php endif; ?>
                   
                   
                     </div><!-- #content -->
                  </div><!-- #container -->
      
               </div>
            </div>
         </div>
      </div>

<?php get_footer(); ?>
