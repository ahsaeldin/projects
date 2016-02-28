<?php

/*

Template Name: Arabic Page

*/

?>



<?php get_header('ar-index'); ?>



	<?php while (have_posts()) : the_post(); ?>	

		<div id="content">	

			<?php the_content(); ?>		

		</div>

	<?php endwhile; ?>

     

<?php get_footer(); ?>