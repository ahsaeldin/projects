function get_header(){
	
    document.write('<div id="logo">');
        document.write('<a href="index.html">');
            document.write('<img src="images/perfect-macro-recoder-logo.png" alt="" title="" border="0" />');
        document.write('</a>');
    document.write('</div>');
	
	add_search_portion();
	
	document.write('<div id="smoothmenu1" class="ddsmoothmenu">');
		document.write('<ul>');
			document.write('<li id="homeli"><a class="top-item" id="home" href="index.html">HOME</a></li>');
			document.write('<li id="more-infoli">');
				document.write('<a class="top-item" id="more-info" href="moreinfo.html">MORE INFO</a>');
				document.write('<ul>');
					document.write('<li><a href="moreinfo.html" id="more-info">Read more</a></li>');
					document.write('<li><a href="screenshots.html">Screenshots</a></li>');
					document.write('<li><a href="video-tour.html">Video tour</a></li>');
					document.write('<li><a href="testimonials.html">Testimonials</a></li>');
					document.write('<li><a class="last-item" href="features.html">Features at a glance</a></li>');
				document.write('</ul>');
			document.write('</li>');
			document.write('<li id="downloadsli"><a class="top-item" id="downloads" href="downloads.html">DOWNLOADS</a></li>');
			document.write('<li id="orderli"><a class="top-item" id="order" href="order.html">ORDER</a></li>');
			document.write('<li id="supportsli">');
				document.write('<a class="top-item" id="support" href="support.html">SUPPORT</a>');
				document.write('<ul>');
					document.write('<li><a href="online-help/online-help.html" target="_blank">Online help</a></li>');
					document.write('<li><a href="faqs.html">FAQs</a></li>');
					document.write('<li><a class="last-item" href="support-request.php">Support request</a></li>');
				document.write('</ul>');
			document.write('</li>');
			document.write('<li id="aboutsli"><a class="top-item" id="about" href="about.html">ABOUT US</a></li>');
		document.write('</ul>');
		document.write('<br style="clear: left" />');
	document.write('</div>');
}     

function add_search_portion(){
	/*
	  1.you need to add both search_script.js & search_style.css and jquery.ajaxLoader.js to the page
	  2.replace #home-content in search_script.js with the div you want to display search results
	  3.insert the search box and search button using this function
	*/
	
	document.write('<div id="page">');
    	document.write('<form id="searchForm" method="post">');
			document.write('<input id="s" type="text" />');
			document.write('<input type="submit" value="" id="submitButton" />');
			document.write('<div id="searchInContainer"></div>');
    	document.write('</form>');
    document.write('</div>');
}

function get_footer(){
	
	document.write('<div id="footer" class="footer-links">');
		document.write('<a class="footer-links" href="index.html">Home</a>');
		document.write(' | ');
		document.write('<a class="footer-links" href="moreinfo.html">More Info</a>');
		document.write(' | ');
		document.write('<a class="footer-links" href="order.html">Order</a>');
		document.write(' | ');
		document.write('<a class="footer-links" href="downloads.html">Downloads</a>');
		document.write(' | ');
		document.write('<a class="footer-links" href="faqs.html">FAQs</a>');
		document.write(' | ');
		document.write('<a class="footer-links" href="about.html">About Us</a>');
		document.write(' | ');
		document.write('<a class="footer-links" href="perfect-macro-recorder-privacy-policy.html">Privacy Policy</a>');
		document.write(' | ');
		document.write('<a class="footer-links" href="terms-and-conditions.html">Terms and Conditions</a>');
		document.write(' | ');
		document.write(' <a class="footer-links" href="support-request.php">Contact Us</a>');
		document.write(' | ');
		document.write(' <a class="footer-links" href="site-map.html">Site Map</a>');
		document.write('<div class="copyright">');
			document.write('Copyright &copy; 2006 - <span id="year"></span>,');/*Inject the year using AJAX*/
			document.write('<a class="footer-links" href="http://www.perfectiontools.com"> Perfection Tools Software</a>.');
			document.write(' All rights reserved. <br />');
		document.write('</div>');
	document.write('</div>');
}

function init_menus(current){
		
	/**************************************************************************************
	* Smooth Navigational Menu- (c) Dynamic Drive DHTML code library (www.dynamicdrive.com)
	* This notice MUST stay intact for legal use
	* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
	***********************************************/
	ddsmoothmenu.init({
	mainmenuid: "smoothmenu1", //menu DIV id
	orientation: 'h', //Horizontal or vertical menu: Set to "h" or "v"
	classname: 'ddsmoothmenu', //class added to menu's outer DIV
	//customtheme: ["#1c5a80", "#18374a"],
	//customtheme: ["#fff", "#fff"],
	contentsource: "markup" //"markup" or ["container_id", "path_to_menu_file"]
	});
	/**************************************************************************************/
	
	$(document).ready(function(){
			
		var site_loaded = $.cookie('site_loaded');		
		if (site_loaded  == 1){
			var NextImageToRender='url(images/why-perfect-macro-recorder'+Math.floor(Math.random()*4)+'.png)';
			//Load Random Image
			$(document).ready(function(){
				//a.saad 10/04/2011: commented to fix a bug of why-perfect-macro-recorder image poz in FF 4
				//$('#green-box').css('background-image',NextImageToRender);
			});
		}
		else{
			$.cookie('site_loaded', '1', { expires: 1 });
		}
		
		//to avoid using loadyear() and wait more time, we just get to point
		$('#year').load('scripts/util.php', "c=time"); 
		//to avoid using disable_right_click and wait more time, we just get to point
		$('.norightclick').bind("contextmenu",function(e){return false;});
		
		/*set the current page before rounding to avoid undo rounding @ IE*/
		$(current).css("background","#9EC040");
		/*display the menu now to avoid displaying menus shapes as rectangles 	before rounding*/
		$('.ddsmoothmenu').css("visibility","visible");
		
		
		//var enable_animation = $.cookie('menu-animation');
		var enable_animation = 0;
		//enable_animation = 0; manually disable animation
				
		if (enable_animation != 0)
		{
			var preHomePos = getRelativePos(document.getElementById('#homeli'));
			var preAboutPos = getRelativePos(document.getElementById('#homeli'));
			var preTopPos = getRelativePos(document.getElementById('#supportsli'));
			var preRightNav = getRelativePos(document.getElementById('#right-nav'));
				
			$('#homeli').css("left","1000px");
			$('#aboutsli').css("left","-1200px");
			$('#supportsli').css("top","1000px");
			$('#more-infoli').css("top","1000px");
			$('#downloadsli').css("top","-600px");
			$('#orderli').css("top","-600px");
		
			
			//Choose easing type form here
			//easeOutElastic easeOutExpo swing easeInOutBounce easeInQuad easeOutQuad easeOutCubic "easeOutSine easeInOutSine"
			
			var general_easing_type = 'swing';
			animate('#homeli','left',preHomePos.offsetLeft,general_easing_type);
			animate('#aboutsli','left',preAboutPos.offsetLeft,general_easing_type);
			animate('#supportsli','top',preTopPos.offsetTop,general_easing_type);
			animate('#more-infoli','top',preTopPos.offsetTop,general_easing_type);
			animate('#downloadsli','top',preTopPos.offsetTop,general_easing_type);			
			animate('#orderli','top',preTopPos.offsetTop,general_easing_type);
			$.cookie('menu-animation', '1', { expires: 1 });
		}
		
	});
	round_corners();
}

function getRelativePos(obj){
	var pos = {offsetLeft:0,offsetTop:0};
	while(obj!==null){
    	pos.offsetLeft += obj.offsetLeft;
    	pos.offsetTop += obj.offsetTop;
    	obj = isIE ? obj.parentElement : obj.offsetParent;
	}
	return pos;
}

function animate(selector,animate_prop,animate_prop_value,easing_type){

	switch (animate_prop)
	{
	case 'top':
  		$(selector).animate({top:animate_prop_value},{duration: 'slow',easing: easing_type});
  		break;
	case 'left':
  		$(selector).animate({left:animate_prop_value},{duration: 'slow',easing: easing_type});
  		break;
	default:
		break;
	}

}

function round_corners(){
	
	$(function(){ 
  	all_corners = {
          tl: { radius: 20 },
          tr: { radius: 20 },
          bl: { radius: 20 },
          br: { radius: 20 },
          antiAlias: true,
          autoPad: true,
          validTags: ["div"]
    };
  	bottom_corners = {
          tl: { radius: 0 },
          tr: { radius: 0 },
          bl: { radius: 20 },
          br: { radius: 20 },
          antiAlias: true,
          autoPad: true,
          validTags: ["div"]
    };
	top_corners = {
          tl: { radius: 20 },
          tr: { radius: 20 },
          bl: { radius: 0 },
          br: { radius: 0 },
          antiAlias: true,
          autoPad: true,
          validTags: ["div"]
    };
	all_rect_corners = {
          tl: { radius: 8 },
          tr: { radius: 8 },
          bl: { radius: 8},
          br: { radius: 8},
          antiAlias: true,
          autoPad: true,
          validTags: ["div"]
    };
	top_rect_corners = {
          tl: { radius: 8 },
          tr: { radius: 8 },
          bl: { radius: 0 },
          br: { radius: 0 },
          antiAlias: true,
          autoPad: true,
          validTags: ["div"]
    };
	
	var curFileName = getCurrentFileName();		
	if( ! $.browser.opera ){
		//a.saad 10/04/2011: commented to fix a bug of why-perfect-macro-recorder image poz in FF 4
		//$('#green-box').corner(all_corners);
	}
	
	$('.top-item').corner(top_corners);
	$('.last-item').corner(bottom_corners);
	
	$('.shadow1').corner(all_rect_corners);
	$('.shadow2').corner(all_rect_corners);
	$('.shadow3').corner(all_rect_corners);
	$('.shadow4').corner(all_rect_corners);		
	$('.top-left-container').corner(all_rect_corners);
	$('.top-right-container').corner(all_rect_corners);
	$('.bottom-left-container').corner(all_rect_corners);
	$('.bottom-right-container').corner(all_rect_corners);
	$('.content-box-container').corner(all_rect_corners);
	$('.top-rect').corner(top_rect_corners);
		
  	});
}

function getCurrentFileName(){
	var url = window.location.pathname;
	var filename = url.substring(url.lastIndexOf('/')+1);
	return filename;
}

function load_year(){
	/*	
	//commented to avoid calling waiting for document.ready and wait more time, we just get to point up in init_menus
	$(document).ready
		(function(){
			$('#year').load('scripts/util.php', "c=time"); 
		}
	);
	*/
}

function get_price(display){
	var price = 11.99;
	if(display === true){
		document.write('$'+price);
	}
	return price;
}

function get_price_per_unit(discount,display){
	var price = get_price(false);
	var rounded_price;
	var price_per_unit = price*(1-discount);
	rounded_price = roundVal(price_per_unit);
	if (display === true){
		document.write(roundVal(price_per_unit) + ' $');
	}
	return rounded_price;
}

function roundVal(val){//return 2 decimal points
	result = val.toFixed(2);
	return result;	
	/*another way
	var dec = 2;
	//round(29.99 * 10^2)=2999
	var result = Math.round(val*Math.pow(10,dec));
	//2999/10^2=29.99
	var result = result/Math.pow(10,dec); 
	return result;
	*/
}

function get_quantity(combo){
	var e = document.getElementById(combo);
	var txt = e.options[e.selectedIndex].text;
	return txt;
}

function get_total_price(discount,combo,display_unit_price,total_price_cell){
	var price_per_unit = get_price_per_unit(discount,display_unit_price);
	if (total_price_cell != undefined){
		var quantity = get_quantity(combo);
		var total_price = roundVal(price_per_unit * quantity);
		document.getElementById(total_price_cell).innerHTML = total_price + ' $';
	}
}

function generate_order_link(combo,order_button){
	var quantity = get_quantity(combo);
	var order_link = 'https://www.plimus.com/jsp/buynow.jsp?contractId=1699584&quantity=' + quantity;
	document.getElementById(order_button).setAttribute('href',order_link);
}

function disable_right_click(selector){
	/*
	//commented to avoid calling waiting for document.ready and wait more time, we just get to point up in init_menus
	$(document).ready(function(){
    	$(selector).bind("contextmenu",function(e){
        return false;
    	});
	});
	*/
}

function get_PMR_Ver(){
	var version = 2.20;
	document.write('v'+roundVal(version));
}

function load_slider_feedback(){
	$(function(){
		$('#contact').contactable({
			subject: 'Message Using Contactable Plugin Ya A7med'
	 	});
	});
}

function scrolldown(idToScrollTo){/*abstraction*/
	if($.browser.opera)
		$('html').animate({scrollTop: $(idToScrollTo).offset().top}, 2000); 
	else 
		$('html,body').animate({scrollTop: $(idToScrollTo).offset().top}, 2000);
}

function scrolldownHomeContent(){
	if($.browser.opera)
		$('html').animate({scrollTop: $('#home-content').offset().top}, 2000); 
	else 
		$('html,body').animate({scrollTop: $('#home-content').offset().top}, 2000);
}