<?php

/*

Template Name: Arabic Index Page

*/

?>



<?php get_header('ar-index'); ?>



	<?php while (have_posts()) : the_post(); ?>	

		<div id="content">	
			<div id="home-page">
            	<?php the_content(); ?>	                
         		<div id="big-image">
					<?php if (function_exists('easing_slider')){ easing_slider(); }; ?>
        		</div>	
                <div id="homeTracker" class="hiddenDiv">home</div><!--hidden text to track home pages -->
            </div>
		</div>

	<?php endwhile; ?>

     

<?php get_footer(); ?>