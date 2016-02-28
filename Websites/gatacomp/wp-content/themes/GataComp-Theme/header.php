<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">	
<html lang="en" xml:lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title><?php bloginfo('name'); ?> <?php if ( is_single() ) { ?> &raquo; Blog Archive <?php } ?> <?php wp_title(); ?></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="keywords" content="civil engineering, Concrete repairs, Protect floors against erosion ,friction and chemicals substance, epoxy work, Positive isolation for the water tanks" />
    <meta name="description" content="Gata - Modern Engineering Company for construction and international trade" />
    <link rel="stylesheet" href="<?php bloginfo('stylesheet_url'); ?>">
    <link type="text/css" rel="stylesheet" href="<?php echo(get_bloginfo('template_directory').'/styles/simpleFAQ.css');?>" />
    <script type="text/javascript" src="<?php echo(get_bloginfo('template_directory').'/js/functions.js');?>" ></script>
    <?php $templatedir =  '\'' . get_bloginfo('template_directory'); ?>
    <?php wp_enqueue_script("jquery"); ?>
    <?php wp_head(); ?> 
</head>
<body>
<div id="main-container">
	<div id="header">
		<img src=<?php echo($templatedir.'/images/ruler.jpg' . '\'');?> width="1000" height="5" alt="" />
			<div id="logo-and-slogan">
				<a href="/home"><img src=<?php echo($templatedir.'/images/main-header-image.jpg' . '\'');?> width="1000" height="145" alt="Gata - Modern Engineering Comp. for construction and international trade" /></a>
			</div>
		<img id="second-ruler" src=<?php echo($templatedir.'/images/ruler.jpg' . '\'');?> width="1000" height="5" alt="" />
		<div id="header-images">        
            <object width="1000" height="145">
            <param name="movie" value=<?php echo($templatedir.'/images/flash-banner.swf' . '\'');?>  />
            <embed src=<?php echo($templatedir.'/images/flash-banner.swf' . '\'');?> type="application/x-shockwave-flash" width="1000" height="145"></embed>
            </object>
		</div>
		<img src=<?php echo($templatedir.'/images/ruler.jpg' . '\'');?> id="third-ruler" height="5"width="1000"  alt="" />
        <div id="menu">
        	<img src=<?php echo($templatedir.'/images/space3.png' . '\'');?> height="24" width="29" alt="" />
            <a href="/services"><img src=<?php echo($templatedir.'/images/services.png' . '\'');?> height="24" width="119" alt="" /></a>
            <a href="/clients"><img src=<?php echo($templatedir.'/images/clients.png' . '\'');?> height="24" width="119" alt="" /></a>
            <a href="/portfolio"><img src=<?php echo($templatedir.'/images/portfolio.png' . '\'');?> height="24" width="119" alt="" /></a>
            <a href="/technical-method"><img src=<?php echo($templatedir.'/images/technical-method.png' . '\'');?> height="24" width="119" alt="" /></a>
            <a href="/about-us"><img src=<?php echo($templatedir.'/images/aboutus.png' . '\'');?> height="24" width="119" alt="" /></a>
            <a href="/home"><img src=<?php echo($templatedir.'/images/home.png' . '\'');?> height="24" width="119" alt="" /></a>          
            <img src=<?php echo($templatedir.'/images/space2.png' . '\'');?> height="24" width="9" alt="" />
            <a href="/contact-us"><img src=<?php echo($templatedir.'/images/contactus.jpg' . '\'');?> height="24" width="82" alt="" /></a>          
            <img src=<?php echo($templatedir.'/images/space3.png' . '\'');?> height="24" width="29" alt="" />    
            <a href="/arabic-home"><img src=<?php echo($templatedir.'/images/arabic.jpg' . '\'');?> height="24" width="30" alt="" /></a>    
            <img id="last-menu-image" src=<?php echo($templatedir.'/images/space4.png' . '\'');?> height="24" width="93" alt="" />
        </div>
</div>