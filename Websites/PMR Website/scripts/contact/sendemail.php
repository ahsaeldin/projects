<?php
// Ajax Contact Form with Validation from jQuery and Anti-bot service from reCaptcha
// By DreamPlus Studio - www.dreamplusstudio.com
// Author: tysoh - www.tysoh.com
// Version: 1.3
// Last updated: 15th of Jun, 2010
//-----------------------------------------------------------------------------------------
//This is the file to send the form to forward the form to email
//together with the public and private key
//You will need to sign up with recaptcha.net to get your own public ans private keys

$mailTo = "support@perfectiontools.com";  // you need to change this to your own email address
$nameFrom = $_POST['emailTo'];
$mailFrom = $_POST['emailFrom'];
$subject = $_POST['subject'];
$message = "Name: ".$_POST['emailTo']."\r\n"."Message: ".$_POST['message'];


$CopyMessageMailTo = $mailFrom;
$CopyMessageMailForm = "support@perfectiontools.com";
$copymessagesubject = "Your e-mail to PerfectionTools Software";
$copymessage = "Here is a copy of the e-mail that you sent to PerfectionTools Software"."\r\n"."\r\n"."-------------- Begin message --------------"."\r\n".
"\r\n".$_POST['message']."\r\n"."\r\n"."-------------- End message --------------"."\r\n"."\r\n"."Thank you very much for your message, we will get back to you soon."."\r\n"."\r\n"."Perfection Tools Software";

mail($mailTo, $subject, $message, "From: ".$mailFrom);
mail($CopyMessageMailTo, $copymessagesubject, $copymessage, "From: ".$CopyMessageMailForm);
?>