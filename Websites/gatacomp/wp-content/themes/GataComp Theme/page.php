<?php
/*
Template Name: English Index Page
*/
?>

<?php get_header();?>
	<?php while (have_posts()) : the_post(); ?>	
		<div id="content">	
			<?php the_content(); ?>	
         <div id="big-image">
		<?php if (function_exists('easing_slider')){ easing_slider(); }; ?>
         </div>
		</div>
	<?php endwhile; ?>
<?php get_footer();?>