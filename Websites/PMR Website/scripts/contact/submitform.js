// -------------------------------------------------------------------
// Ajax Contact Form with Validation from jQuery and Anti-bot service from reCaptcha
// By DreamPlus Studio - www.dreamplusstudio.com
// Author: tysoh - www.tysoh.com
// Version: 1.3
// Last updated: 15th of Jun, 2010
// -------------------------------------------------------------------
// Purpose of this file: To validate the user input
// -------------------------------------------------------------------

$(document).ready(function(){
	$("#submit").click(function(){					   				   
		$(".error").hide();
		var hasError = false;
		var emailReg = /^([\w-\.]+@([\w-]+\.)+[\w-]{2,4})?$/;
		
		var emailToVal = $("#nameFrom").val();
		if(emailToVal == '') {
			$("#nameFrom").after('<div class="error">Please enter your name.</div>');
			hasError = true;
		}
		
		var emailFromVal = $("#emailFrom").val();
		if(emailFromVal == '') {
			$("#emailFrom").after('<div class="error">Please enter your email address.</div>');
			hasError = true;
		} else if(!emailReg.test(emailFromVal)) {	
			$("#emailFrom").after('<div class="error">Please enter a valid email address..</div>');
			hasError = true;
		}
		
		var subjectVal = $("#subject").val();
		if(subjectVal == '') {
			$("#subject").after('<div class="error">Please enter the subject.</div>');
			hasError = true;
		}
		
		var messageVal = $("#message").val();
		if(messageVal == '') {
			$("#message").after('<div class="error">Please enter the message.</div>');
			hasError = true;
		}
		
			var recaptchaVal = $("#recaptcha_response_field").val();
		if(recaptchaVal == '') {
			$("#recaptcha_response_field").after('<div class="error">Please validate the Form.</div>');
			hasError = true;
			
		}	

		
		if(recaptchaVal != '')  {
			var challengeField = $("#recaptcha_challenge_field").val();
    		var responseField = $("#recaptcha_response_field").val();
   			 var html = $.ajax({
   			 type: "POST",
   			 url: "scripts/contact/ajax.recaptcha.php",
   			 data: "recaptcha_challenge_field=" + challengeField + "&recaptcha_response_field=" + responseField,
   			 async: false
   			 }).responseText;
   
    		if(html == "success"){}
   			else {
      		 $("#recaptcha_response_field").after('<div class="error">Validation error. Please try again.</div>');
      		  Recaptcha.reload();
              hasError = true;
  			 }
		}
		
		if(hasError == false) {
			$(this).hide();
			$.post("scripts/contact/sendemail.php",
   				{ emailTo: emailToVal, emailFrom: emailFromVal, subject: subjectVal, message: messageVal },
   					function(data){
						$("#sendEmail").slideUp("slow");
						$("#sendEmail").after('<div class="thankyou">Thank you.</div><br />Your email has been sent and a copy of your email has been sent to your own email address <br />as a confirmation.<br /><br />Please note that all inquiries are personally processed and as such it may take up to 48 hours <br />to get a response.<br /><br />');				
                                       
   					}
				 );
		}
		
		return false;
	});						   
});