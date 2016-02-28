/** 
 * Copyright CSSJockey | http://www.cssjockey.com
 */
var $cj = jQuery.noConflict();
$cj(document).ready(function(){
	$cj('.fade').animate( { borderRightWidth:"1px" }, 1000).fadeOut(500);

	$cj(".cjcontent").hide(0);
	$cj("h1.cjhead").click(function(){
		var id = $cj(this).attr("id");
		$cj("."+id).slideToggle(100);
		$cj(this).toggleClass("cjminus");
	})


// Toggle Category Ids
	$cj(".cjcats").hide(0);
	$cj("#cjshowcatids").click(function(){
		$cj(".cjcats").slideToggle();
	})
	
// Toggle Page Ids
	$cj(".cjpages").hide(0);
	$cj("#cjshowpageids").click(function(){
		$cj(".cjpages").slideToggle();
	})

// Date Checks

    $cj("input[name*='cjsave']").click(function(){
        var spdate = $cj("input[name*='cj_splash_page_launch_day']").val();
        var spmonth = $cj("input[name*='cj_splash_page_launch_month']").val();
        var spyear = $cj("input[name*='cj_splash_page_launch_year']").val();
            if(spdate < 1 || spdate > 31){
                alert('Enter a valid day of the month');
                return false;
            }else if(spmonth < 1 || spmonth > 12){
                alert('Enter a valid month');
                return false;
            }else if((spyear % 4) == 0 && spmonth == 02 && spdate > 29){
                alert("It's February of a Leap Year :)");
                return false;
            }else if((spyear % 4) != 0 && spmonth == 02 && spdate > 28){
                alert("It's February :)");
                return false;
            }
    })




})