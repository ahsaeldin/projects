<?php
/**
 * Template Name: Support
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
         <div class="support-posts">

				<?php $loop = new WP_Query(array('post_type' => 'support-post', 'posts_per_page' => 6)); ?>
				<?php if ($loop->have_posts()): ?>
                  <ul>
							<?php $posts_counter = 0; ?>
								<?php while ($loop->have_posts()) :	$loop->the_post(); $posts_counter++; ?>
                           <li class="question-<?php echo $posts_counter; ?>">
                              <div>
                                 <i><a href="<?php the_permalink(); ?>"><?php the_title(''); ?></a></i>
                                 <?php the_excerpt(''); ?>
                                 <b><a href="<?php the_permalink(); ?>">Read More</a></b>
                              </div>
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
            
                  <?php get_sidebar('newsletter'); ?>
            
                  <div id="container">
                     <div id="content" role="main">
            				
                        <div class="help-block">
									<?php $wp_query = new WP_Query(array('post_type' => 'help-center', 'posts_per_page' => 1)); ?>
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
                                    	<div class="wrapper">
                                       	<?php the_excerpt( __( '<span>Go to Online Help Center</span>', 'theme1144' ) ); ?>
                                       </div>
                                    </div><!-- .entry-content -->
                                    
                                    
                                 </div><!-- #post-## -->
                              
                            <?php endwhile; ?>
                        </div>
                     
                     </div><!-- #content -->
                  </div><!-- #container -->
               
               </div>
            </div>
         </div>
      </div>

<?php get_footer(); ?>
