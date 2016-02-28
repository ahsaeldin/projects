<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="en" xml:lang="en" xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>Perfect Macro Recorder Support Request Form - Perfection Tools Software</title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
   		<meta name="keywords" content="macro recorder,keyboard and mouse macro,keyboard macro,keystroke macro,keyboard macro recorder,macro tool" />
		<meta name="description" content="Support request form for Perfect Macro Recorder" /> 
        <link rel="stylesheet" type="text/css" href="style/style.css" media="screen" />
		<link rel="stylesheet" type="text/css" href="style/ddsmoothmenu.css" />
        <link rel="stylesheet" type="text/css" href="style/support-request-form.css" media="screen" />
        <link rel="stylesheet" type="text/css" href="style/search_style.css" />
		<script type="text/javascript" src="scripts/func.js" ></script>
	    <script type="text/javascript" >
			var RecaptchaOptions = {
   			theme : 'custom',
   			tabindex : 2,
   			custom_theme_widget: 'recaptcha_widget'
			};
		</script>
    </head>
    <body>
		<div id="main-container">
			<div id="header">
        		<script type="text/javascript">get_header('');</script> 
    		</div>    
            <div id="green-box"><!--start of greenbox -->
                <div class="clock"> 
    				<img src="images/Perfect-Macro-Recorder-Screenshot.jpg" alt="Screensoft Of Perfect Macro Recorder" title="" />
				</div>
    			<div id="right-nav">
                <a href="PerfectMacroRecorder.exe" onclick="DoThis();"><img src="images/download-macro-recorder-now.gif" border="0" alt="Download Perfect Macro Recorder" /></a>
    			</div>
            </div><!--end of greenbox -->
    		<div id="main-content">
    			<div id="home-content">
                	<div id="home-content-left">
                        <div id="middle">
                            <form action="support-request.php" method="post" id="sendEmail">
                                <div class="orange"><h2 id="srf">Support Request Form</h2></div>
                                <p>
                                Useful help resources that may answer your questions:<br />
                                </p>
                                <ul>
                                    <li><a href="online-help.html">Perfect Macro Recorder Online Help</a></li>
                                    <li><a href="video-tour.html">Video Tour</a></li>
                                    <li><a href="features.html">Features</a></li>
                                    <li><a href="faqs.html">FAQs</a></li>
                                </ul>
                                <p>
                                But If you still cannot find an answer for your questions, we value your inquiry a lot, so please do not <br />hesitate to contact us if you have any further questions about Perfect Macro Recorder.
                                <br /><br /> 
                                </p>
                                <p class="alert">* All fields are required</p>
                                    <table width="583" border="0" cellpadding="0" cellspacing="0" class="contacttable">
                                          <tr>
                                            <td width="138" valign="top"><strong>*Name:</strong></td>
                                            <td width="345"><input type="text" name="nameFrom" id="nameFrom" value="<?= $_POST['nameFrom']; ?>"  class="textfield" /></td>
                                          </tr>
                                          <tr>
                                            <td valign="top"><strong>*Email:</strong></td>
                                            <td><input type="text" name="emailFrom" id="emailFrom" value="<?= $_POST['emailFrom']; ?>" class="textfield" /></td>
                                          </tr>
                                          <tr>
                                            <td valign="top"><strong>*Subject: </strong></td>
                                            <td><input type="text" name="subject" id="subject" value="<?= $_POST['subject']; ?>" class="textfield" /></td>
                                          </tr>
                                          <tr>
                                            <td valign="top"><strong>*Message: </strong></td>
                                            <td><textarea name="message" id="message" rows="10" cols="40"><?= $_POST['message']; ?></textarea></td>
                                          </tr>
                                          <tr>
                                            <td valign="top">&nbsp;</td>
                                            <td></td>
                                          </tr>
                                          <tr>
                                            <td valign="top">&nbsp;</td>
                                            <td><div id="recaptcha_widget" style="display:none">
                                                <div id="recaptcha_image"></div>
                                                <div class="recaptcha_only_if_incorrect_sol" style="color:red">Incorrect please try again</div>
                                                <span class="recaptcha_only_if_image">Please validate the texts above:</span> <br />
                                                <input type="text" id="recaptcha_response_field" name="recaptcha_response_field" size="25" style="height:20px;border:1px;border-color:#b5b5b5;border-style:solid;" />
                                                <?= $_POST['recaptcha_response_field']; ?>
                                                <div><a href="javascript:Recaptcha.reload()">Get another CAPTCHA</a></div>
                                                <div class="recaptcha_only_if_image"></div>
                                                <div class="recaptcha_only_if_audio"><a href="javascript:Recaptcha.switch_type('image')">Get an image CAPTCHA</a></div>
                                              </div>
                                              <?php
                                                require_once('scripts/contact/recaptchalib.php');
                                    
                                                // Get a key from http://recaptcha.net/api/getkey
                                                $publickey = "6Ldi57wSAAAAABk1JD7Tk-oRmhzjZ7cyyCgyAXfG"; // you need to change this to your own key
                                    
                                                # the response from reCAPTCHA
                                                $resp = null;
                                                # the error code from reCAPTCHA, if any
                                                $error = null;
                                    
                                                echo recaptcha_get_html($publickey, $error);
                                                ?></td>
                                          </tr>
                                          <tr>
                                            <td valign="top">&nbsp;</td>
                                            <td>&nbsp;</td>
                                          </tr>
                                          <tr>
                                            <td>&nbsp;</td>
                                            <td><input type="submit" value="Submit" name="submit" id="submit" />
                                                <input name="reset" type="reset" class="fb" /></td>
                                          </tr>
                                    </table>
                            </form>
                        </div>     	
                     </div>
                     <div id="home-content-right">
                        <div>
                            <img class="center" src="images/perfect-macro-recorder-support-request.jpg" height="200" width="200" 
                            alt="Contact US &amp; Support Form" />
                        </div>
                 	</div>
                </div>
    		</div><!--end of main content-->
    		<script type="text/javascript">get_footer();</script>
		</div> <!--end of main container-->
        <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.4.2/jquery.min.js"></script>
        <script type="text/javascript" src="scripts/contact/submitform.js" ></script>
        <script type="text/javascript" src="scripts/jquery.easing.1.3.js"></script>
        <script type="text/javascript" src="scripts/jquery.cookie.js"></script>
        <script type="text/JavaScript" src="scripts/jquery.curvycorners.source.js"></script>
        <script type="text/javascript" src="scripts/ddsmoothmenu.js"></script>
        <script type="text/javascript" src="scripts/search_script.js"></script>
        <script type="text/javascript" src="scripts/jquery.ajaxLoader.js" ></script>
		<script type="text/javascript">
			init_menus();
		</script>
        <!-- Start of StatCounter Code -->
        <script type="text/javascript">
        	var sc_project=6057275; 
            var sc_invisible=1; 
            var sc_security="47c7a0ed"; 
        </script>
        <script type="text/javascript" src="http://www.statcounter.com/counter/counter_xhtml.js"></script>
        <noscript>
        <div class="statcounter">
        	<a title="visit tracker on tumblr" class="statcounter" href="http://www.statcounter.com/tumblr/"><img class="statcounter" width="1" height="1" src="http://c.statcounter.com/6057275/0/47c7a0ed/1/" alt="visit tracker on tumblr" /></a>
        </div>
        </noscript>
    	<!-- End of StatCounter Code -->
	</body>
</html>
