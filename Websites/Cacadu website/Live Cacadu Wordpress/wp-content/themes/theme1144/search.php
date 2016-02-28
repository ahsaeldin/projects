<?php
/**
 * The template for displaying Search Results pages.
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
            
                  <?php get_sidebar(search); ?>

                  <div id="container">
                     <div id="content" role="main">
            
            <?php if ( have_posts() ) : ?>
                        <h1 class="page-title"><?php printf( __( 'Search Results for: %s', 'theme1144' ), '<span>' . get_search_query() . '</span>' ); ?></h1>
                        <?php
                        /* Run the loop for the search to output the results.
                         * If you want to overload this in a child theme then include a file
                         * called loop-search.php and that will be used instead.
                         */
                         get_template_part( 'loop', 'search' );
                        ?>
            <?php else : ?>
                        <div id="post-0" class="post no-results not-found">
                           <h2 class="entry-title"><?php _e( 'Nothing Found', 'theme1144' ); ?></h2>
                           <div class="entry-content">
                              <p><?php _e( 'Sorry, but nothing matched your search criteria. Please try again with some different keywords.', 'theme1144' ); ?></p>
                              <?php get_search_form(); ?>
                           </div><!-- .entry-content -->
                        </div><!-- #post-0 -->
            <?php endif; ?>
                     </div><!-- #content -->
                  </div><!-- #container -->
      
               </div>
            </div>
         </div>
      </div>

<?php get_footer(); ?>
