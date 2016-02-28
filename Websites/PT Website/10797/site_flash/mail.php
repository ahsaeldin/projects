<?php
$Name = $_POST["Nam"];
$Add = $_POST["EM"];
$Mess = $_POST["Me"];
If ($Name != "" And $Add != "" And $Mess != "")
{
$Final = "From: $Name ($Add)\n $Mess";

//ini_set functions are only needed if you do not want the default settings
//SMTP server default is localhost

//ini_set("SMTP", "[smtp server address]");

//SMTP port is 25

//ini_set("smtp_port", [port number]);
mail("cprinahmed@yahoo.com", "[Subject]", $Final);
echo '<font face="Arial" size="2">Thanks for your email, we will get back to you shortly.<br>To return to the homepage click <a href="."><span style="text-decoration: none">here</span></a></font><br><br><br><br><br><br><br><br><br><br><br><br><br>';
}
else
{
echo '<font face="Arial" size="2">You must fill in all the feilds please go <a href="javascript:history.back()"><span style="text-decoration: none">back</span></a></font><br><br><br><br><br><br><br><br><br><br><br><br><br><br>';
}
?>