<?php
/*
Plugin Name: Custom Comment
Plugin URI: http://imaprogrammer.wordpress.com/2011/01/11/custom-comment/
Description: Customize your comment form with more fields
Version: 2.1.6
Author: Moein Akbarof
Author URI: http://imaprogrammer.wordpress.com/
License: GLPv2
*/
?>
<?php
function CComment_init() {
	// embed the javascript file that makes the AJAX request
    wp_enqueue_script( 'CComment-ajax-request', plugin_dir_url( __FILE__ ) . 'js/CComment.js', array( 'jquery' ) );         
 
	wp_localize_script( 'CComment-ajax-request', 'CComment_ajax_var', array(
	// URL to wp-admin/admin-ajax.php to process the request
	'ajaxurl'          => admin_url( 'admin-ajax.php' ),
 
	// generate a nonce with a unique ID "CComment_ajax_nonce"
	// so that you can check it later when an AJAX request is sent
	'nonce' => wp_create_nonce( 'CComment-ajax-nonce' ),
	)
	);
}    
add_action('init', 'CComment_init');
add_action( 'wp_ajax_CComment-submit', 'CComment_ajax_submit' );
//add_action('wp_ajax_my_action', 'CComent_ajax_respond');
add_action('comment_form_after_fields', 'CComment_form');
add_action('comment_form_logged_in_after', 'CComment_form');
add_action('comment_post', 'CComment_comment_post');
add_action('admin_menu', 'CComment_modify_admin_menu');
add_action('delete_comment', 'CComment_delete_comment');
add_filter('comment_author', 'CComment_link');
//add_action('check_comment_flood', 'check_required_CComment_fields');
register_activation_hook (__FILE__, 'set_CComment_options');

//Check for the required extra fields
function check_required_CComment_fields() {
	$custom_fields = array_filter(explode(",", get_option('CComment_fields')));
	$error = array();
	foreach ($custom_fields as $field_name) {
		$custom_field_arr = explode(";CuCo;", get_option($field_name));
		if ($custom_field_arr[2] == 1 && $custom_field_arr[1] == 1 && $_POST['CuCo_'.$field_name] == "")
			$error[] = $custom_field_arr[0];
	}
	if(count($error) != 0){
		$error_desc = 'Error: please fill the required fields ('.implode(",", $error).').';
		wp_die($error_desc);
	}

}

function CComment_ajax_submit() {
	$nonce = $_POST["nonce"];
	 if ( ! wp_verify_nonce( $nonce, 'CComment-ajax-nonce' ))
		die('Not verified');
	if (!current_user_can('manage_options'))
		die("You don't have the permission");
	
	if(isset($_POST["CC_delete"])){
		$field_name = $_POST["CC_name"];
		$custom_fields = array_filter(explode(",", get_option('CComment_fields')));
		$all_comments = array_filter(explode(",", get_option('CComment_comments_id')));
		//Search through the fields array to find the requested field
		foreach ($custom_fields as $key => $custom_field)
			if($custom_field == $field_name){
				//Remove the requested field values from all comments meta
				foreach( $all_comments as $comment_id)
					delete_comment_meta($comment_id, "CuCo_".$field_name);
				//Remove the request field
				unset($custom_fields[$key]);
				//Remove the requested field's data
				delete_option($field_name);
			}
		update_option('CComment_fields', implode(",", $custom_fields));
		die("1"); 
	}
	elseif(isset($_POST["CC_edit"])){
		$field_name = $_POST["CC_name"];
		$desc = $_POST["new_desc"];
		$option = get_option($field_name, "CC_NOT_EXIST");
		if ($option == "CC_NOT_EXIST")
			die("This field doesn't exit!");
		$custom_field_arr = explode(";CuCo;", $option);
		$act = $custom_field_arr[1];
		$desc .= ";CuCo;".$act;
		update_option($field_name, $desc);
		die("1");
	}
	elseif(isset($_POST["CC_add"])){
		if (strpos(get_option('CComment_fields'),$_POST["CC_name"]) !== false )
			die("The name already Exists!");
		$option = get_option($_POST["CC_name"], "NOT_EXIST");
		if ($option != "NOT_EXIST")
			if (strpos($option, ";CuCo;") === false)
				die("The name already Exists!");
		update_option($_POST["CC_name"], $_POST["CC_desc"].";CuCo;1;CuCo;0");
		$custom_fields = array_filter(explode(",", get_option('CComment_fields')));
		$custom_fields[] = $_POST["CC_name"];
		update_option('CComment_fields', implode(",",$custom_fields));
		die("1");
	}
	elseif(isset($_POST["CC_act"])){
		$option = get_option($_POST["CC_name"], "NOT_EXIST");
		if($option == "NOT_EXIST")
			die("The name doesn't exist!");
		$custom_field_arr = explode(";CuCo;", $option);
		$custom_field_arr[1] = $_POST["CC_act"];
		update_option($_POST["CC_name"], implode(";CuCo;", $custom_field_arr));
		die("1");
	}
	elseif(isset($_POST["CC_req"])){
		$option = get_option($_POST["CC_name"], "NOT_EXIST");
		if($option == "NOT_EXIST")
			die("The name doesn't exist!");
		$custom_field_arr = explode(";CuCo;", $option);
		$custom_field_arr[2] = $_POST["CC_req"];
		update_option($_POST["CC_name"], implode(";CuCo;", $custom_field_arr));
		die("1");
	}
	exit();
}

//Print the extra fields in the comment form
function CComment_form() {
	$custom_fields = array_filter(explode(",", get_option('CComment_fields')));
	foreach ($custom_fields as $custom_field) {
		$custom_field_arr = explode(";CuCo;", get_option($custom_field));
		if($custom_field_arr[2] == 1) 
			$required = '<span class="required"> *</span>';
		else
			$required = '';
		if ($custom_field_arr[1] == 1)
			echo '<p><label for="'.$custom_field.'">'.$custom_field_arr[0].'</label>'.$required.'<input type="text" size="30" value="" name="CuCo_'.$custom_field.'" id="CuCo_'.$custom_field.'"></p>';
	}
}

//Save the custom fields for the comment
function CComment_comment_post($comment_id) {
	//Check required files
	check_required_CComment_fields();
	
	$custom_fields = array_filter(explode(",", get_option('CComment_fields')));
	foreach ($custom_fields as $custom_field)
		if($_POST["CuCo_".$custom_field] != "")
			add_comment_meta($comment_id, 'CuCo_'.$custom_field, $_POST['CuCo_'.$custom_field]);
	$CC_comments_id = array_filter(explode(",", get_option('CComment_comments_id')));
	$CC_comments_id[] = $comment_id;
	update_option('CComment_comments_id', implode(",", $CC_comments_id));
}

//Add the info link in front of the author's name
function CComment_link($author) {
	if(is_admin())
		$author .= ' <a href="options-general.php?page='.basename(__FILE__).'&c='.get_comment_ID().'" target="_blank">Info</a>';
	return $author;
}


//Set the options when the plugin is activated
function set_CComment_options() {
	//Holds the name of the extra fields
	add_option('CComment_fields', '');
	//Holds the id of posted comments with extra fields data
	add_option('CComment_comments_id','');
}

//Remove comment id from the list of comments(the list will be used for uninstalling the plugin
function CComment_delete_comment($comment_id) {
	$CC_comments_id = array_filter(explode(",", get_option('CComment_comments_id')));
	foreach ($CC_comments_id as $key => $CC_comment_id)
		if($CC_comment_id == $comment_id)
			unset($CC_comments_id[$key]);
}

////////////////This part for the option page
//Layout of the option page
function CComment_option_page() {
	echo '<div class="wrap">
			<h2>Custom Comments Fields</h2><hr />';
	if (isset($_GET["c"])) {
		$CComment_fields = array_filter(explode(",", get_option('CComment_fields')));
		echo '
			<table class="widefat comments fixed" id="CComment_table">
				<tr>
					<th width="40%">Description</th>
					<th width="60%">Value</th>
				</tr>';
		foreach ($CComment_fields as $CComment_field){
			$CComment_field_arr = explode(";CuCo;", get_option($CComment_field));
			$CComment_field_desc = $CComment_field_arr[0];
			$CComment_field_active = $CComment_field_arr[1];
			$Comment_field_value = get_comment_meta($_GET["c"], "CuCo_".$CComment_field);
			if ($CComment_field_active == 1)
				echo '<tr><td>'.$CComment_field_desc.'</td><td>'.$Comment_field_value[0].'</td></tr>';
		}
		echo '</table>';
	}
	else
		print_CComment_fields_form();
}

//Prints the setting table of extra fields
function print_CComment_fields_form() {
echo '
	<form method="post" id="CComment_form">
		<fieldset class="options">
			<table class="widefat comments fixed" id="CComment_table">
			<tr>
				<th width="15%">Name</th>
				<th width="45%">Description</th>
				<th width="10%">Edit</th>
				<th width="10%">Active</th>
				<th width="10%">Required</th>
				<th width="10%">Delete</th>
			</tr>';
				$CComment_fields = array_filter(explode(",", get_option('CComment_fields')));
				foreach ($CComment_fields as $CComment_field) {
					$CComment_field_arr = explode(";CuCo;", get_option($CComment_field));
					$CComment_field_desc = $CComment_field_arr[0];
					$CComment_field_active = $CComment_field_arr[1];
					$CComment_field_required = $CComment_field_arr[2];
					if ($CComment_field_active == 1){
						$active_button = '<input type="button" value="Deactivate" class="cc_act" title="Click to deactivate this field" />';
						$color = "#FFFFFF";
					}
					else{
						$active_button = '<input type="button" value="Activate" class="cc_act" title="Click to activate this field" />';
						$color = "#EEEEEE";
					}
					$checked_required = '';
					if ($CComment_field_required == 1)
						$required_button = '<input type="button" value="No" class="cc_req" title="Click to make this field not required" />';
					else
						$required_button = '<input type="button" value="Yes" class="cc_req" title="Click to make this field required" />';
					echo"
						<tr style=\"background-color:$color;\">
							<td>
								$CComment_field
							</td>
							<td>
								$CComment_field_desc
							</td>
							<td>
								<input type=\"button\" value=\"Edit\" name=\"del\" class=\"cc_edit\" title=\"Click to edit this field\" />
							</td>
							<td>
								$active_button
							</td>
							<td>
								$required_button
							</td>
							<td>
								<div><input type=\"button\" value=\"Delete\" name=\"del\" class=\"cc_delete\" title=\"Click to delete this field\" /></div>
							</td>
						</tr>";
				}
			
		echo	'</table>
			<p>
				The name should be unique and can only contain [a-z A-Z 0-9 space and  _]<br />
				<label for="CComment_name">Name: </label>
				<input type="text" name="CComment_name" id="CComment_name" value="" />
				<label for="CComment_desc">Description: </label>
				<input type="text" name="CComment_desc" id="CComment_desc" value="" />
				<input type="submit" name="add" id="add" value="Add" />
			</p>
			<h3>To support this plugin please <a href="http://wordpress.org/extend/plugins/customcomment/" target="_blank">rate</a> and consider <a href="http://imaprogrammer.wordpress.com/donation/" target="_blank">donating</a></h3>
			<h3>Got any feedback or comment? please let me know <a href="http://imaprogrammer.wordpress.com/2011/01/11/custom-comment/" target="_blank">Here</a></h3>
			<h3><a href="http://imaprogrammer.wordpress.com/2011/01/07/tiprotector/" target="_blank">Protect your precious posts from theives!</a>
			<h2 style="text-align: right">Powered By <a href="http://www.imaprogrammer.wordpress.com" target="_blank">Moein</a></h2>
			
		</fieldset>
	</form>
</div>
';
}

//The function for modifying the admin menu
function CComment_modify_admin_menu() {
	add_options_page('Custom Comment', 'Custom Comment', 'manage_options', basename(__FILE__), 'CComment_option_page');
}

?>
