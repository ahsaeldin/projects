<?php

// Ajax Contact Form with Validation from jQuery and Anti-bot service from reCaptcha
// By DreamPlus Studio - www.dreamplusstudio.com
// Author: tysoh - www.tysoh.com
// Version: 1.3
// Last updated: 15th of Jun, 2010
//-----------------------------------------------------------------------------------------
//This is the file to verify the user input with the given captcha
//together with the public and private key
//You will need to sign up with recaptcha.net to get your own public ans private keys


require_once('recaptchalib.php');
$publickey = "6Ldi57wSAAAAABk1JD7Tk-oRmhzjZ7cyyCgyAXfG"; // you got this from the signup page
$privatekey = "6Ldi57wSAAAAAGi4BGaRSB7OvjmWYwhG4RWkL1Z7";

$resp = recaptcha_check_answer ($privatekey,
                                $_SERVER["REMOTE_ADDR"],
                                $_POST["recaptcha_challenge_field"],
                                $_POST["recaptcha_response_field"]);

if ($resp->is_valid) {
    ?>success<?
}
else 
{
  ?>unsuccess<?
}
?>
