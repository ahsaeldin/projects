<?php
/**
 * options-page in wp-admin
 */

// export options
if ( isset( $_GET['_mw_adminimize_export'] ) ) {
	_mw_adminimize_export();
	die();
}

//include( 'adminimize_admin_bar.php' );

function _mw_adminimize_options() {
	global $wpdb, $_wp_admin_css_colors, $wp_version, $wp_roles, $table_prefix;

	$_mw_adminimize_user_info = '';

	//get array with userroles
	$user_roles = get_all_user_roles();
	$user_roles_names = get_all_user_roles_names();

	// update options
	if ( ( isset($_POST['_mw_adminimize_action']) && $_POST['_mw_adminimize_action'] == '_mw_adminimize_insert') && $_POST['_mw_adminimize_save'] ) {

		if ( function_exists('current_user_can') && current_user_can('manage_options') ) {
			check_admin_referer('mw_adminimize_nonce');

			_mw_adminimize_update();

		} else {
			$myErrors = new _mw_adminimize_message_class();
			$myErrors = '<div id="message" class="error"><p>' . $myErrors->get_error('_mw_adminimize_access_denied') . '</p></div>';
			wp_die($myErrors);
		}
	}
	
	// import options
	if ( ( isset($_POST['_mw_adminimize_action']) && $_POST['_mw_adminimize_action'] == '_mw_adminimize_import') && $_POST['_mw_adminimize_save'] ) {

		if ( function_exists('current_user_can') && current_user_can('manage_options') ) {
			check_admin_referer('mw_adminimize_nonce');
			
			_mw_adminimize_import();
			
		} else {
			$myErrors = new _mw_adminimize_message_class();
			$myErrors = '<div id="message" class="error"><p>' . $myErrors->get_error('_mw_adminimize_access_denied') . '</p></div>';
			wp_die($myErrors);
		}
	}
	
	// deinstall options
	if ( ( isset($_POST['_mw_adminimize_action']) && $_POST['_mw_adminimize_action'] == '_mw_adminimize_deinstall') &&  ($_POST['_mw_adminimize_deinstall_yes'] != '_mw_adminimize_deinstall') ) {

		$myErrors = new _mw_adminimize_message_class();
		$myErrors = '<div id="message" class="error"><p>' . $myErrors->get_error('_mw_adminimize_deinstall_yes') . '</p></div>';
		wp_die($myErrors);
	}

	if ( ( isset($_POST['_mw_adminimize_action']) && $_POST['_mw_adminimize_action'] == '_mw_adminimize_deinstall') && $_POST['_mw_adminimize_deinstall'] && ($_POST['_mw_adminimize_deinstall_yes'] == '_mw_adminimize_deinstall') ) {

		if ( function_exists('current_user_can') && current_user_can('manage_options') ) {
			check_admin_referer('mw_adminimize_nonce');

			_mw_adminimize_deinstall();

			$myErrors = new _mw_adminimize_message_class();
			$myErrors = '<div id="message" class="updated fade"><p>' . $myErrors->get_error('_mw_adminimize_deinstall') . '</p></div>';
			echo $myErrors;
		} else {
			$myErrors = new _mw_adminimize_message_class();
			$myErrors = '<div id="message" class="error"><p>' . $myErrors->get_error('_mw_adminimize_access_denied') . '</p></div>';
			wp_die($myErrors);
		}
	}
	
	// load theme user data
	if ( ( isset($_POST['_mw_adminimize_action']) && $_POST['_mw_adminimize_action'] == '_mw_adminimize_load_theme') && $_POST['_mw_adminimize_load'] ) {
		if ( function_exists('current_user_can') && current_user_can('edit_users') ) {
			check_admin_referer('mw_adminimize_nonce');
			
			$myErrors = new _mw_adminimize_message_class();
			$myErrors = '<div id="message" class="updated fade"><p>' . $myErrors->get_error('_mw_adminimize_load_theme') . '</p></div>';
			echo $myErrors;
		} else {
			$myErrors = new _mw_adminimize_message_class();
			$myErrors = '<div id="message" class="error"><p>' . $myErrors->get_error('_mw_adminimize_access_denied') . '</p></div>';
			wp_die($myErrors);
		}
	}
	
	if ( ( isset($_POST['_mw_adminimize_action']) && $_POST['_mw_adminimize_action'] == '_mw_adminimize_set_theme') && $_POST['_mw_adminimize_save'] ) {
		if ( function_exists('current_user_can') && current_user_can('edit_users') ) {
			check_admin_referer('mw_adminimize_nonce');
			
			_mw_adminimize_set_theme();
			
			$myErrors = new _mw_adminimize_message_class();
			$myErrors = '<div id="message" class="updated fade"><p>' . $myErrors->get_error('_mw_adminimize_set_theme') . '</p></div>';
			echo $myErrors;
		} else {
			$myErrors = new _mw_adminimize_message_class();
			$myErrors = '<div id="message" class="error"><p>' . $myErrors->get_error('_mw_adminimize_access_denied') . '</p></div>';
			wp_die($myErrors);
		}
	}
?>
	<div class="wrap">
		<?php 
		// Backend Options for all roles
		require_once('inc-options/minimenu.php');
		?>
		
		<form name="backend_option" method="post" id="_mw_adminimize_options" action="?page=<?php echo $_GET['page'];?>" >
		<?php 
		// Backend Options for all roles
		require_once('inc-options/backend_options.php');
		
		// global options on all pages in backend for diffferent roles
		require_once('inc-options/global_options.php');
		
		// Menu Submenu Options
		require_once('inc-options/menu_options.php');
		
		// Write Page Options
		require_once('inc-options/write_post_options.php');
		
		// Write Page Options
		require_once('inc-options/write_page_options.php');
		
		// Links Options 
		require_once('inc-options/links_options.php');

		// WP Nav Menu Options 
		require_once('inc-options/wp_nav_menu_options.php');
		?>
		</form>
		
		<div id="poststuff" class="ui-sortable meta-box-sortables">
			<div class="postbox">
				<div class="handlediv" title="<?php _e('Click to toggle'); ?>"><br/></div>
				<h3 class="hndle" id="set_theme"><?php _e('Set Theme', FB_ADMINIMIZE_TEXTDOMAIN ) ?></h3>
				<div class="inside">
					<br class="clear" />
					
					<?php if (  !isset($_POST['_mw_adminimize_action']) || !($_POST['_mw_adminimize_action'] == '_mw_adminimize_load_theme') ) { ?>
					<form name="set_theme" method="post" id="_mw_adminimize_set_theme" action="?page=<?php echo $_GET['page'];?>" >
							<?php wp_nonce_field('mw_adminimize_nonce'); ?>
							<p><?php _e('For better peformance with many users on your blog; load only userlist, when you will change the theme options for users.', FB_ADMINIMIZE_TEXTDOMAIN ); ?></p>
							<p id="submitbutton">
								<input type="hidden" name="_mw_adminimize_action" value="_mw_adminimize_load_theme" />
								<input type="submit" name="_mw_adminimize_load" value="<?php _e('Load User Data', FB_ADMINIMIZE_TEXTDOMAIN ); ?> &raquo;" class="button button-primary" />
							</p>
					</form>
					<?php }
					if (  isset($_POST['_mw_adminimize_action']) && ($_POST['_mw_adminimize_action'] == '_mw_adminimize_load_theme') ) { ?>
						<form name="set_theme" method="post" id="_mw_adminimize_set_theme" action="?page=<?php echo $_GET['page'];?>" >
							<?php wp_nonce_field('mw_adminimize_nonce'); ?>
							<table class="widefat">
								<thead>
									<tr class="thead">
										<th>&nbsp;</th>
										<th class="num"><?php _e('User-ID') ?></th>
										<th><?php _e('Username') ?></th>
										<th><?php _e('Display name publicly as') ?></th>
										<th><?php _e('Admin-Color Scheme') ?></th>
										<th><?php _e('User Level') ?></th>
										<th><?php _e('Role') ?></th>
									</tr>
								</thead>
								<tbody id="users" class="list:user user-list">
									<?php
									$wp_user_search = $wpdb->get_results("SELECT ID, user_login, display_name FROM $wpdb->users ORDER BY ID");
	
									$style = '';
									foreach ( $wp_user_search as $userid ) {
										$user_id       = (int) $userid->ID;
										$user_login    = stripslashes($userid->user_login);
										$display_name  = stripslashes($userid->display_name);
										$current_color = get_user_option('admin_color', $user_id);
										$user_level    = (int) get_user_option($table_prefix . 'user_level', $user_id);
										$user_object   = new WP_User($user_id);
										$roles         = $user_object->roles;
										$role          = array_shift($roles);
										if ( function_exists('translate_user_role') )
											$role_name   = translate_user_role( $wp_roles->role_names[$role] );
										elseif ( function_exists('before_last_bar') )
											$role_name   = before_last_bar( $wp_roles->role_names[$role], 'User role' );
										else
											$role_name   = strrpos( $wp_roles->role_names[$role], '|' );
										
										$style = ( ' class="alternate"' == $style ) ? '' : ' class="alternate"';
										$return  = '';
										$return .= '<tr>' . "\n";
										$return .= "\t" . '<td><input type="checkbox" name="mw_adminimize_theme_items[]" value="' . $user_id . '" /></td>' . "\n";
										$return .= "\t" . '<td class="num">'. $user_id .'</td>' . "\n";
										$return .= "\t" . '<td>'. $user_login .'</td>' . "\n";
										$return .= "\t" . '<td>'. $display_name .'</td>' . "\n";
										$return .= "\t" . '<td>'. $current_color . '</td>' . "\n";
										$return .= "\t" . '<td class="num">'. $user_level . '</td>' . "\n";
										$return .= "\t" . '<td>'. $role_name . '</td>' . "\n";
										$return .= '</tr>' . "\n";
	
										echo $return;
									}
									?>
										<tr valign="top">
											<td>&nbsp;</td>
											<td>&nbsp;</td>
											<td>&nbsp;</td>
											<td>&nbsp;</td>
											<td>
												<select name="_mw_adminimize_set_theme">
													<?php foreach ( $_wp_admin_css_colors as $color => $color_info ): ?>
														<option value="<?php echo $color; ?>"><?php echo $color_info->name . ' (' . $color . ')' ?></option>
													<?php endforeach; ?>
													</select>
											</td>
											<td>&nbsp;</td>
											<td>&nbsp;</td>
										</tr>
								</tbody>
							</table>
							<p id="submitbutton">
								<input type="hidden" name="_mw_adminimize_action" value="_mw_adminimize_set_theme" />
								<input type="hidden" name="_mw_adminimize_load" value="_mw_adminimize_load_theme" />
								<input type="submit" name="_mw_adminimize_save" value="<?php _e('Set Theme', FB_ADMINIMIZE_TEXTDOMAIN ); ?> &raquo;" class="button button-primary" />
							</p>
						</form>
					<?php } ?>
					
					<p><a class="alignright button" href="javascript:void(0);" onclick="window.scrollTo(0,0);" style="margin:3px 0 0 30px;"><?php _e('scroll to top', FB_ADMINIMIZE_TEXTDOMAIN); ?></a><br class="clear" /></p>
				</div>
			</div>
		</div>
		
		<div id="poststuff" class="ui-sortable meta-box-sortables">
			<div class="postbox">
				<div class="handlediv" title="<?php _e('Click to toggle'); ?>"><br/></div>
				<h3 class="hndle" id="import"><?php _e('Export/Import Options', FB_ADMINIMIZE_TEXTDOMAIN ) ?></h3>
				<div class="inside">
					<br class="clear" />
					
					<h4><?php _e('Export', FB_ADMINIMIZE_TEXTDOMAIN ) ?></h4>
					<form name="export_options" method="get" action="">
						<p><?php _e('You can save a .seq file with your options.', FB_ADMINIMIZE_TEXTDOMAIN ) ?></p>
						<p id="submitbutton">
							<input type="hidden" name="_mw_adminimize_export" value="true" />
							<input type="submit" name="_mw_adminimize_save" value="<?php _e('Export &raquo;', FB_ADMINIMIZE_TEXTDOMAIN ) ?>" class="button" />
						</p>
					</form>
					
					<h4><?php _e('Import', FB_ADMINIMIZE_TEXTDOMAIN ) ?></h4>
					<form name="import_options" enctype="multipart/form-data" method="post" action="?page=<?php echo $_GET['page'];?>">
						<?php wp_nonce_field('mw_adminimize_nonce'); ?> 
						<p><?php _e('Choose a Adminimize (<em>.seq</em>) file to upload, then click <em>Upload file and import</em>.', FB_ADMINIMIZE_TEXTDOMAIN ) ?></p>
						<p>
							<label for="datei_id"><?php _e('Choose a file from your computer', FB_ADMINIMIZE_TEXTDOMAIN ) ?>: </label>
							<input name="datei" id="datei_id" type="file" />
						</p>
						<p id="submitbutton">
							<input type="hidden" name="_mw_adminimize_action" value="_mw_adminimize_import" />
							<input type="submit" name="_mw_adminimize_save" value="<?php _e('Upload file and import &raquo;', FB_ADMINIMIZE_TEXTDOMAIN ) ?>" class="button" />
						</p>
					</form>
					<p><a class="alignright button" href="javascript:void(0);" onclick="window.scrollTo(0,0);" style="margin:3px 0 0 30px;"><?php _e('scroll to top', FB_ADMINIMIZE_TEXTDOMAIN); ?></a><br class="clear" /></p>
					
				</div>
			</div>
		</div>

		<div id="poststuff" class="ui-sortable meta-box-sortables">
			<div class="postbox">
				<div class="handlediv" title="<?php _e('Click to toggle'); ?>"><br/></div>
				<h3 class="hndle" id="uninstall"><?php _e('Deinstall Options', FB_ADMINIMIZE_TEXTDOMAIN ) ?></h3>
				<div class="inside">

					<p><?php _e('Use this option for clean your database from all entries of this plugin. When you deactivate the plugin, the deinstall of the plugin <strong>clean not</strong> all entries in the database.', FB_ADMINIMIZE_TEXTDOMAIN ); ?></p>
					<form name="deinstall_options" method="post" id="_mw_adminimize_options_deinstall" action="?page=<?php echo $_GET['page'];?>">
						<?php wp_nonce_field('mw_adminimize_nonce'); ?>
						<p id="submitbutton">
							<input type="submit" name="_mw_adminimize_deinstall" value="<?php _e('Delete Options', FB_ADMINIMIZE_TEXTDOMAIN ); ?> &raquo;" class="button-secondary" />
							<input type="checkbox" name="_mw_adminimize_deinstall_yes" value="_mw_adminimize_deinstall" />
							<input type="hidden" name="_mw_adminimize_action" value="_mw_adminimize_deinstall" />
						</p>
					</form>
					<p><a class="alignright button" href="javascript:void(0);" onclick="window.scrollTo(0,0);" style="margin:3px 0 0 30px;"><?php _e('scroll to top', FB_ADMINIMIZE_TEXTDOMAIN); ?></a><br class="clear" /></p>

				</div>
			</div>
		</div>

		<div id="poststuff" class="ui-sortable meta-box-sortables">
			<div class="postbox" >
				<div class="handlediv" title="<?php _e('Click to toggle'); ?>"><br/></div>
				<h3 class="hndle" id="about"><?php _e('About the plugin', FB_ADMINIMIZE_TEXTDOMAIN ) ?></h3>
				<div class="inside">
				
					<p><?php _e('Further information: Visit the <a href="http://bueltge.de/wordpress-admin-theme-adminimize/674/">plugin homepage</a> for further information or to grab the latest version of this plugin.', FB_ADMINIMIZE_TEXTDOMAIN); ?></p>
					<p>
					<span style="float: left;">
						<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
						<input type="hidden" name="cmd" value="_s-xclick">
						<input type="hidden" name="hosted_button_id" value="4578111">
						<input type="image" src="https://www.paypal.com/en_US/i/btn/btn_donate_SM.gif" border="0" name="submit" alt="<?php _e('PayPal - The safer, easier way to pay online!', FB_ADMINIMIZE_TEXTDOMAIN); ?>">
						<img alt="" border="0" src="https://www.paypal.com/de_DE/i/scr/pixel.gif" width="1" height="1">
					</form>
					</span>
					<?php _e('You want to thank me? Visit my <a href="http://bueltge.de/wunschliste/">wishlist</a> or donate.', FB_ADMINIMIZE_TEXTDOMAIN); ?>
					</p>
					<p>&copy; Copyright 2008 - <?php echo date('Y'); ?> <a href="http://bueltge.de">Frank B&uuml;ltge</a></p>
					<p class="textright" style="color:#ccc"><small><?php echo $wpdb->num_queries; ?>q, <?php timer_stop(1); ?>s</small></p>
					<p><a class="alignright button" href="javascript:void(0);" onclick="window.scrollTo(0,0);" style="margin:3px 0 0 30px;"><?php _e('scroll to top', FB_ADMINIMIZE_TEXTDOMAIN); ?></a><br class="clear" /></p>
					
				</div>
			</div>
		</div>

		<script type="text/javascript">
		<!--
		<?php if ( version_compare( $wp_version, '2.7alpha', '<' ) ) { ?>
		jQuery('.postbox h3').prepend('<a class="togbox">+</a> ');
		<?php } ?>
		jQuery('.postbox h3').click( function() { jQuery(jQuery(this).parent().get(0)).toggleClass('closed'); } );
		jQuery('.postbox .handlediv').click( function() { jQuery(jQuery(this).parent().get(0)).toggleClass('closed'); } );
		jQuery('.postbox.close-me').each(function() {
			jQuery(this).addClass("closed");
		});
		//-->
		</script>

	</div>
<?php
}
?>
