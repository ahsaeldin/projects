<?php

// Zopim Customize Widget Page

function zopim_customize_widget() {
	global $current_user;
	$ul        = $current_user->data->first_name;
	$useremail = $current_user->data->user_email;
	$greetings = json_to_array(get_option('zopimGreetings'));
	$message   = "";

	if (count($_POST) > 0) {
		update_option('zopimLang',         $_POST["zopimLang"]);
		update_option('zopimPosition',     $_POST["zopimPosition"]);
		update_option("zopimBubbleEnable", $_POST["zopimBubbleEnable"] );
		update_option('zopimColor',        $_POST["zopimColor"]);
		update_option('zopimTheme',        $_POST["zopimTheme"]);
		update_option('zopimBubbleTitle',  stripslashes($_POST["zopimBubbleTitle"]));
		update_option('zopimBubbleText',   stripslashes($_POST["zopimBubbleText"]));

		update_checkbox("zopimGetVisitorInfo");
		update_checkbox("zopimHideOnOffline");
		update_checkbox("zopimUseGreetings");
		update_checkbox("zopimUseBubble");

		if (isset($_POST['zopimUseGreetings']) && $_POST['zopimUseGreetings'] != "") {
			$greetings->online->window  = stripslashes($_POST["zopimOnlineLong"]);
			$greetings->online->bar     = stripslashes($_POST["zopimOnlineShort"]);
			$greetings->away->window    = stripslashes($_POST["zopimAwayLong"]);
			$greetings->away->bar       = stripslashes($_POST["zopimAwayShort"]);
			$greetings->offline->window = stripslashes($_POST["zopimOfflineLong"]);
			$greetings->offline->bar    = stripslashes($_POST["zopimOfflineShort"]);

			update_option('zopimGreetings', to_json($greetings));
		}

		$message = "<b>Changes saved!</b><br>";
	}

	zopimme();

	$accountDetails = getAccountDetails(get_option('zopimSalt'));

	if (get_option('zopimCode')=="zopim") {
		$message = '<div class="metabox-holder">
	<div class="postbox">
		<h3 class="hndle"><span>Customizing in Demo Mode</span></h3>
		<div style="padding:10px;line-height:17px;">
			Currently customizing in demo mode. Messages in this widget will go to Zopim staff. The chat widget will not appear on your site until you <a href="admin.php?page=zopim_account_config">activate / link up an account</a>. <br>
		</div>
	</div>
		</div>';
		$accountDetails->widget_customization_enabled = 1;
		$accountDetails->color_customization_enabled = 1;
	}
	else if (isset($accountDetails->error)) {
		$message = '<div class="metabox-holder">
								<div class="postbox">
								<h3 class="hndle"><span>Account no longer linked!</span></h3>
								<div style="padding:10px;line-height:17px;">
									We could not connect to your Zopim account. As a result, this customization page is running in demo mode.<br> Please <a href="admin.php?page=zopim_account_config">check your password in account setup</a> and try again.
								</div>
								</div>
								</div>';
	}
	else {
		$message .= "Click 'Save Changes' when you're done. Happy customizing!";
	}
	// unset($accountDetails->widget_customization_enabled);
	// unset($accountDetails->color_customization_enabled);
?>

	<script type="text/javascript">

	function updateWidget() {
		var lang = document.getElementById('zopimLang').options[ document.getElementById('zopimLang').options.selectedIndex ].value;
		$zopim.livechat.setLanguage(lang);

		if (document.getElementById("zopimGetVisitorInfo").checked) {
			$zopim.livechat.setName('<?php echo $ul; ?>');
			$zopim.livechat.setEmail('<?php echo $useremail; ?>');
		}
		else {
			$zopim.livechat.setName('Visitor');
			$zopim.livechat.setEmail('');
		}

		document.getElementById("zopimHideOnOffline").checked? $zopim.livechat.button.setHideWhenOffline(true): $zopim.livechat.button.setHideWhenOffline(false);

		$zopim.livechat.window.setColor(document.getElementById("zopimColor").value);
		$zopim.livechat.window.setTheme(document.getElementById("zopimTheme").value);

		if (document.getElementById("zopimUseBubble").checked) {
			$zopim.livechat.bubble.setTitle(document.getElementById("zopimBubbleTitle").value);
			$zopim.livechat.bubble.setText(document.getElementById("zopimBubbleText").value);
		}
		else {
			$zopim.livechat.bubble.setTitle('Questions?');
			$zopim.livechat.bubble.setText('Click here to chat with us!');
		}

		$zopim.livechat.setGreetings({
			'online':  [document.getElementById("zopimOnlineShort").value, document.getElementById("zopimOnlineLong").value],
			'offline': [document.getElementById("zopimOfflineShort").value, document.getElementById("zopimOfflineLong").value],
			'away':    [document.getElementById("zopimAwayShort").value, document.getElementById("zopimAwayLong").value]
		});
	}

	function updatePosition() {
		var position = document.getElementById('zopimPosition').options[ document.getElementById('zopimPosition').options.selectedIndex ].value;
		$zopim.livechat.button.setPosition(position);
	}

	function updateBubbleStatus() {
		var value = document.getElementById("zopimBubbleEnable").value;
		switch (value) {
			case 'default':
				$zopim.livechat.bubble.reset();
				break;
			case 'show':
				$zopim.livechat.bubble.show();
				break;
			case 'hide':
				$zopim.livechat.bubble.hide();
				break;
		}
	}

	function greetingsChanged() {
		var inputs     = [ 'zopimOnlineShort', 'zopimAwayShort', 'zopimOfflineShort',
											 'zopimOnlineLong' , 'zopimAwayLong', 'zopimOfflineLong' ];
		var isDisabled = false;

		document.getElementById('zopimUseGreetings').checked? (isDisabled = false): (isDisabled = true);

		for (var i=0; i<inputs.length; i++) {
			document.getElementById(inputs[i]).disabled = isDisabled;
		}

		updateWidget();
	}

	function bubbleChanged() {
		var inputs     = [ 'zopimBubbleTitle', 'zopimBubbleText' ];
		var isDisabled = false;

		document.getElementById('zopimUseBubble').checked? (isDisabled = false): (isDisabled = true);

		for (var i = 0; i < inputs.length; i++) {
			document.getElementById(inputs[i]).disabled = isDisabled;
		}

		updateWidget();
	}

	var timer;
	function updateSoon() {
		clearTimeout(timer);
		timer = setTimeout("updateWidget()", 300);
	}

	</script>

<style type="text/css">
.smallExplanation {
background:#FAFAFA;
color:#667788;
font-size:8pt;
line-height:13px;
margin:4px 0 0 0;
padding:8px;
display: inline-block;
}
.inputtextshort {
width:200px;
}
.inputtext {
width:450px;
}
.secthead {
border-bottom:1px solid #EEEEEE;
color:#8899AA;
font-size:13px;
line-height:21px;
}
.sethead {
	width:200px;
}
.swatch {
	float: left;
	width: 15px
}
.swatch:hover {
	background-image:url(http://www.zopim.com/static/images/colorselectbg.gif);
	cursor:pointer;
}
.sorry {
	color:#c33;
}
</style>

<div class="wrap">
<div id="icon-themes" class="icon32"><br/></div><h2>Customize your widget</h2>

<?php echo $message; ?>
<form method="post" action="admin.php?page=zopim_customize_widget">
<div class="metabox-holder">
	<div class="postbox">
		<h3 class="hndle"><span>General Settings</span></h3>
		<div style="padding:10px;">
		<table class="form-table">
			<tr valign="top">
				<th scope="row" class="sethead">Language</th>
				<td>
					<select name="zopimLang" id="zopimLang" onchange="updateWidget()">
<?php
	$languages = get_languages();
	echo generate_options($languages, get_option('zopimLang'));
?>
				</select>
				</td>
				</tr>
				<tr valign="top" style="display:none;">
				<th scope="row">Use Logged in Username / Email</th>
				<td><input onchange="updateWidget()" type="checkbox" id="zopimGetVisitorInfo" name="zopimGetVisitorInfo" value="zopimGetVisitorInfo" <?php if (get_option('zopimGetVisitorInfo')!="disabled") { echo "checked='checked'"; } ?> /></td>
				</tr>
				<tr valign="top">
				<th scope="row" class="sethead">Position</th>
				<td>

				<select name="zopimPosition" id="zopimPosition" onchange="updatePosition()">
<?php
	 $positions = array("br" => "Bottom Right", "bl" => "Bottom Left");
	 echo generate_options($positions, get_option('zopimPosition'));
?>
				</select>
				</td>
				</tr>

				<tr valign="top">
				<th scope="row">Hide chat bar when offline<br>
				<!-- <div class="smallExplanation">Hide the chat bar when no agents are available to answer questions. This prevents visitors from sending you offline messages. </div>  -->
				</th>
					<td>
						<input onchange="updateWidget()" type="checkbox" id="zopimHideOnOffline" name="zopimHideOnOffline" value="zopimHideOnOffline" <?php if (get_option('zopimHideOnOffline') && get_option('zopimHideOnOffline')!="disabled") { echo "checked='checked'"; } ?> /> Prevents visitors from sending you offline messages
						<a href="#" onclick="$zopim.livechat.button.show();return false">(Show preview chat bar)</a>
					</td>
				</tr>
		</table>
		</div>
	</div>
</div>
<div class="metabox-holder">
	<div class="postbox">
		<h3 class="hndle"><span>Color & Theme Settings</span></h3>
		<div style="padding:10px;">
		Theme settings show up instantly on the preview widget. Try it out!<br/>
		<table class="form-table" style="width: 700px">
				<tr valign="top">
				<td colspan="2">
				<input type="hidden" id="zopimColor" name="zopimColor" value="<?php echo get_option('zopimColor'); ?>">
<?php

	 if ($accountDetails->color_customization_enabled == 1) {
			echo "<div style='display:inline-block;border:11px solid #888;background:#888;color:#fee;'>";
			$colors = curl_get_url(ZOPIM_COLORS_LIST);
			$colors = explode("\n", $colors);

			$i=0;
			foreach ($colors as $color) {
				 echo "<div class='swatch' style='background-color: $color;' onclick=\"document.getElementById('zopimColor').value='$color'; updateWidget();\">&nbsp</div>";
				 if (++$i%40==0) {
						echo "<br>";
				 }
			}
			echo "<br><a href=# style='color:#ff8' onclick=\"document.getElementById('zopimColor').value=''; updateWidget();\">Restore default color</a></div>";
	 } else {
			echo "<div class='sorry'>Sorry, your plan does not allow for color customization. Please upgrade to enjoy choice of color!</div>";
	 }
?>
				</td>
				</tr>
				<tr valign="top">
				<th scope="row" class="sethead">Select A Theme</th>
				<td style="width: 400px"><div align="left">
<?php

	if ($accountDetails->widget_customization_enabled == 1) {
		echo '<select name="zopimTheme" id="zopimTheme" onchange="updateWidget()">';
		$themes = curl_get_url(ZOPIM_THEMES_LIST);
		$themes = valuekeys(explode("\n", $themes));
		ksort($themes);

		echo generate_options($themes, get_option('zopimTheme'));
		echo "</select> <a href='#' onclick='\$zopim.livechat.window.toggle();return false;'>Open the chat panel</a> to see changes";
	}
	else {
		echo "<div class='sorry'>Sorry, your plan does not allow for theme customization. Please upgrade to enjoy choice of themes!</div>";
		echo '<input type=hidden value="" name="zopimTheme" id="zopimTheme">';
	 }
?>
				</td>
				</tr>
		</table>
		</div>
	</div>
</div>

<?php
	(get_option('zopimUseBubble') && get_option('zopimUseBubble')!="disabled")? ($useBubble = true): ($useBubble = false);
?>
<div class="metabox-holder">
	<div class="postbox">
		<h3 class="hndle"><span>Help Bubble Settings</span></h3>
		<div style="padding:10px;">
		<table class="form-table">
			<tr valign="top">
				<th scope="row">Display Help Bubble<br></th>
				<td><select onchange="updateBubbleStatus()" type="checkbox" id="zopimBubbleEnable" name="zopimBubbleEnable"> <?php $bubble_modes = array("default" => "Let user decide", "show" => "Always show", "hide" => "Always hide"); echo generate_options($bubble_modes, get_option('zopimBubbleEnable')); ?>	</select>
			</tr>
			<tr>
				<td colspan="2"><div class="secthead">Help Bubble Text</div></td>
			</tr>
			<tr>
				<td colspan="2"><input type="checkbox" id="zopimUseBubble" onchange="bubbleChanged()" value="zopimUseBubble" name="zopimUseBubble" <?php if ($useBubble) { echo "checked='checked'"; } ?>/> Enable help bubble text settings (override dashboard settings) - save settings to view changes.</td>
			</tr>
			<tr valign="top">
				<th scope="row" class="sethead">Help Bubble Title</th>
				<td><input class="inputtextshort" name="zopimBubbleTitle" id="zopimBubbleTitle" onKeyup="updateSoon()" value="<?php echo get_option('zopimBubbleTitle'); ?>"  <?php if(!$useBubble) { echo "disabled='disabled'"; } ?>><!-- <a href="#" onclick="updateBubbleStatus();">Refresh</a></td> -->
			</tr>
			<tr valign="top">
				<th scope="row" class="sethead">Help Bubble Message</th>
				<td><input class="inputtext" name="zopimBubbleText" id="zopimBubbleText" onKeyup="updateSoon()" value="<?php echo get_option('zopimBubbleText'); ?>"  <?php if(!$useBubble) { echo "disabled='disabled'"; } ?>></td>
			</tr>
		</table>
		</div>
	</div>
</div>

<?php
	(get_option('zopimUseGreetings') && get_option('zopimUseGreetings')!="disabled")? ($useGreetings = true): ($useGreetings = false);
?>
<div class="metabox-holder">
	<div class="postbox">
		<h3 class="hndle"><span>Greeting Message Settings</span></h3>
		<div style="padding:10px;">
		<table class="form-table">
			<tr>
				<td colspan="2"><input type="checkbox" id="zopimUseGreetings" onchange="greetingsChanged()" value="zopimUseGreetings" name="zopimUseGreetings" <?php if ($useGreetings) { echo "checked='checked'"; } ?>/>	Enable greeting message settings (override dashboard settings) </td>
			</tr>
			<tr>
				<td colspan="2"><div class="secthead">Message Shown on Chat Bar (max 26 characters)</div></td>
			</tr>
			<tr valign="top">
				<th scope="row" class="sethead">Online</th>
				<td><input class="inputtextshort" name="zopimOnlineShort" id="zopimOnlineShort" onKeyup="updateSoon()" value="<?php echo $greetings->online->bar; ?>" <?php if (!$useGreetings) echo "disabled='disabled'"?> maxlength="26"></td>
			</tr>
			<tr valign="top">
				<th scope="row" class="sethead">Away</th>
				<td><input class="inputtextshort" name="zopimAwayShort" id="zopimAwayShort" onKeyup="updateSoon()"  value="<?php echo $greetings->away->bar; ?>"  <?php if (!$useGreetings) echo "disabled='disabled'"?>  maxlength="26"></td>
			</tr>
			<tr valign="top">
				<th scope="row" class="sethead">Offline</th>
				<td><input class="inputtextshort" name="zopimOfflineShort" id="zopimOfflineShort" onKeyup="updateSoon()" value="<?php echo $greetings->offline->bar; ?>" <?php if (!$useGreetings) echo "disabled='disabled'"?>  maxlength="26"></td>
			</tr>
			<tr>
				<td colspan="2"><div class="secthead">Message Shown on Chat Panel (max 140 characters)</div></td>
			</tr>
			<tr valign="top">
				<th scope="row" class="sethead">Online</th>
				<td><textarea class="inputtext" name="zopimOnlineLong" id="zopimOnlineLong" onKeyup="updateSoon()"  <?php if (!$useGreetings) echo "disabled='disabled'"?> ><?php echo $greetings->online->window; ?></textarea></td>
			</tr>
			<tr valign="top">
				<th scope="row" class="sethead">Away</th>
				<td><textarea class="inputtext" name="zopimAwayLong" id="zopimAwayLong" onKeyup="updateSoon()" <?php if (!$useGreetings) echo "disabled='disabled'"?> ><?php echo $greetings->away->window; ?></textarea></td>
			</tr>
			<tr valign="top">
				<th scope="row" class="sethead">Offline</th>
				<td><textarea class="inputtext" name="zopimOfflineLong" id="zopimOfflineLong" onKeyup="updateSoon()"  <?php if (!$useGreetings) echo "disabled='disabled'"?> ><?php echo $greetings->offline->window; ?></textarea></td>
			</tr>
		</table>
		</div>
	</div>
</div>

<p class="submit">
	<input type="submit" class="button-primary" value="<?php _e('Save Changes') ?>" />
</p>

</form>
</div>

<?php }

function valuekeys($array) {
	 $newarray = array();
	 foreach ($array as $s) { $newarray[$s] = $s; }
	 return $newarray;
}

function generate_options($options, $current) {
	 $out = "";
	 foreach ($options as $key => $value) {
			if ($value != "") {
				 $isselected = "";
				 if ($current == $key) {
						$isselected = "selected";
				 }
				 $out .= '<option value="'.$key.'" '.$isselected.'>'.$value.'</option>';
			}
	 }
	 return $out;
}

function get_languages() {
		$langjson = '{"--":" - Auto Detect - ","ar":"Arabic","bn":"Bengali","bg":"Bulgarian","zh_CN":"Chinese (China)","zh_TW":"Chinese (Taiwan)","hr":"Croatian","cs":"Czech","da":"Danish","nl":"Dutch; Flemish","et":"Estonian","fo":"Faroese","fi":"Finnish","fr":"French","ka":"Georgian","de":"German","el":"Greek","he":"Hebrew","hu":"Hungarian","is":"Icelandic","id":"Indonesian","it":"Italian","ja":"Japanese","ko":"Korean","ku":"Kurdish","lv":"Latvian","lt":"Lithuanian","mk":"Macedonian","ms":"Malay","nb":"Norwegian Bokmal","fa":"Persian","pl":"Polish","pt":"Portuguese","pt_BR":"Portuguese (Brazil)","ro":"Romanian","ru":"Russian","sr":"Serbian","sk":"Slovak","sl":"Slovenian","es":"Spanish; Castilian","sv":"Swedish","th":"Thai","tr":"Turkish","uk":"Ukrainian","ur":"Urdu","vi":"Vietnamese"}';
	 return json_to_array($langjson);
}

function update_checkbox($fieldname) {
	 if (isset($_POST["$fieldname"]) && $_POST["$fieldname"] != "")
			update_option($fieldname, $_POST["$fieldname"]);
	 else
			update_option($fieldname, "disabled");
}

?>
