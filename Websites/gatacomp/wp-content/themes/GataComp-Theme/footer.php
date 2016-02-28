     <div id="footer">

     	<div id="developedby">

        	Developed by <span id="shadowname"><a href="http://www.mooparmghor.com">Ahmed Saad</a></span>

        </div>

     </div>
</div>

<?php 
	if (is_page('Home') || is_page('الصفحة الرئيسية') )
	{
		$custom_field_keys = get_post_custom_keys();
		foreach ( $custom_field_keys as $key => $value ) {
			$valuet = trim($value);
			if ( '_' == $valuet{0} )
				continue;
				$mykey_values = get_post_custom_values($value);
				if($value == 'comments'){
					$c = 0;
					foreach ( $mykey_values as $key => $key_value ) 
					{
						$c = $c + 1;
						echo ('<div id="hiddenDiv'.$c.'" '.'class="hiddenDiv">'.$key_value.'</div>'); 
					};
				};
		}
	}
?>
    

<script type="text/javascript" src="<?php echo(get_bloginfo('template_directory').'/js/jquery.simpleFAQ-0.7.min.js');?>" ></script>

</head>

<script type="text/javascript">

	var $js = jQuery.noConflict();

	var currentPageName = GetPageName();

	DisplayPortoflio()

	//RotateHomePageComments();

    /*Functions*/
	function DisplayPortoflio(){
		$js(document).ready(function () {
			$js('ul#faqList').simpleFAQ();			  
		});
	}

	function RotateHomePageComments(){
		$js(document).ready(function () {
			if(document.getElementById('homeTracker').innerHTML == 'home' ){
			var count = 1;/*escape first image*/ 	
			setInterval(function() {
				 count = count + 1;
				 divVar = "#hiddenDiv"+count;
				 //$js("#big-image-comment-text").slideUp();
				 $js("#big-image-comment-text").html($js(".nivo-caption").html());
				// $js("#big-image-comment-text").slideDown(2000);
				 
				 if (count == 9)
					count = 0; 
			}, 5000);
		}
		});
	}
	
	function GetPageName(){
			var url = window.location.pathname;
			var urlLength = url.length;
			var lastSlash = url.lastIndexOf('/');
			if (urlLength - lastSlash <= 1)
				url = url.substring(0,lastSlash);
			
			var pageName = url.substring(url.lastIndexOf('/') + 1);  
			return pageName;
	}
		
</script>

<?php wp_footer();?>

</body>

</html>