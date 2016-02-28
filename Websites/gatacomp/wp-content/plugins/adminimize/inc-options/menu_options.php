<?php
// options for menu, submenu
?>

		<div id="poststuff" class="ui-sortable meta-box-sortables">
			<div class="postbox">
				<div class="handlediv" title="<?php _e('Click to toggle'); ?>"><br/></div>
				<h3 class="hndle" id="config_menu"><?php _e('Menu Options', FB_ADMINIMIZE_TEXTDOMAIN ); ?></h3>
				<div class="inside">
					<br class="clear" />
					
					<table summary="config_menu" class="widefat">
						<thead>
							<tr>
								<th><?php _e('Menu options - Menu, <span style=\"font-weight: 400;\">Submenu</span>', FB_ADMINIMIZE_TEXTDOMAIN ); ?></th>

								<?php foreach ($user_roles_names as $role_name) { ?>
										<th><?php _e('Deactivate for', FB_ADMINIMIZE_TEXTDOMAIN ); echo '<br/>' . $role_name; ?></th>
								<?php } ?>

							</tr>
						</thead>
						<tbody>
							<?php
							$wp_menu    = _mw_adminimize_getOptionValue('mw_adminimize_default_menu');
							$wp_submenu = _mw_adminimize_getOptionValue('mw_adminimize_default_submenu');
							
							if ( empty($wp_menu) ) {
								global $menu;
								
								$wp_menu = $menu;
							}
							if ( !isset($wp_submenu) ) {
								global $submenu;
								
								$wp_submenu = $submenu;
							}
							
							foreach ($user_roles as $role) {
								$disabled_metaboxes_post_[$role]  = _mw_adminimize_getOptionValue('mw_adminimize_disabled_metaboxes_post_'. $role .'_items');
								$disabled_metaboxes_page_[$role]  = _mw_adminimize_getOptionValue('mw_adminimize_disabled_metaboxes_page_'. $role .'_items');
							}

							$metaboxes = array(
								'#contextual-help-link-wrap',
								'#screen-options-link-wrap',
								'#titlediv',
								'#pageslugdiv',
								'#tagsdiv,#tagsdivsb,#tagsdiv-post_tag',
								'#formatdiv',
								'#categorydiv,#categorydivsb',
								'#category-add-toggle',
								'#postexcerpt',
								'#trackbacksdiv',
								'#postcustom',
								'#commentsdiv',
								'#passworddiv',
								'#authordiv',
								'#revisionsdiv',
								'.side-info',
								'#notice',
								'#post-body h2',
								'#media-buttons',
								'#wp-word-count',
								'#slugdiv,#edit-slug-box',
								'#misc-publishing-actions',
								'#commentstatusdiv',
								'#editor-toolbar #edButtonHTML, #quicktags'
							);

							if ( function_exists('current_theme_supports') && current_theme_supports( 'post-thumbnails', 'post' ) )
								array_push($metaboxes, '#postimagediv');
							if (class_exists('SimpleTagsAdmin'))
								array_push($metaboxes, '#suggestedtags');
							if (function_exists('tc_post'))
								array_push($metaboxes, '#textcontroldiv');
							if (class_exists('HTMLSpecialCharactersHelper'))
								array_push($metaboxes, '#htmlspecialchars');
							if (class_exists('All_in_One_SEO_Pack'))
								array_push($metaboxes, '#postaiosp, #aiosp');
							if (function_exists('tdomf_edit_post_panel_admin_head'))
								array_push($metaboxes, '#tdomf');
							if (function_exists('post_notification_form'))
								array_push($metaboxes, '#post_notification');
							if (function_exists('sticky_add_meta_box'))
								array_push($metaboxes, '#poststickystatusdiv');

							$metaboxes_names = array(
								__('Help'),
								__('Screen Options'),
								__('Title', FB_ADMINIMIZE_TEXTDOMAIN),
								__('Permalink', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Tags', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Format'),
								__('Categories', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Add New Category', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Excerpt', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Trackbacks', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Custom Fields'),
								__('Comments', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Password Protect This Post', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Post Author'),
								__('Post Revisions'),
								__('Related, Shortcuts', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Messages', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('h2: Advanced Options', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Media Buttons (all)', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Word count', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Post Slug'),
								__('Publish Actions', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Discussion'),
								__('HTML Editor Button')
							);
							
							if ( function_exists('current_theme_supports') && current_theme_supports( 'post-thumbnails', 'post' ) )
								array_push($metaboxes_names, __('Post Thumbnail') );
							if (class_exists('SimpleTagsAdmin'))
								array_push($metaboxes_names, __('Suggested tags from'));
							if (function_exists('tc_post'))
								array_push($metaboxes_names, __('Text Control'));
							if (class_exists('HTMLSpecialCharactersHelper'))
								array_push($metaboxes_names, __('HTML Special Characters'));
							if (class_exists('All_in_One_SEO_Pack'))
								array_push($metaboxes_names, __('All in One SEO Pack'));
							if (function_exists('tdomf_edit_post_panel_admin_head'))
								array_push($metaboxes_names, 'TDOMF');
							if (function_exists('post_notification_form'))
								array_push($metaboxes_names, 'Post Notification');
							if (function_exists('sticky_add_meta_box'))
								array_push($metaboxes, 'Post Sticky Status');
							
							// add own post options
							$_mw_adminimize_own_post_values  = _mw_adminimize_getOptionValue('_mw_adminimize_own_post_values');
							$_mw_adminimize_own_post_values = preg_split( "/\r\n/", $_mw_adminimize_own_post_values );
							foreach ( (array) $_mw_adminimize_own_post_values as $key => $_mw_adminimize_own_post_value ) {
								$_mw_adminimize_own_post_value = trim($_mw_adminimize_own_post_value);
								array_push($metaboxes, $_mw_adminimize_own_post_value);
							}
							
							$_mw_adminimize_own_post_options = _mw_adminimize_getOptionValue('_mw_adminimize_own_post_options');
							$_mw_adminimize_own_post_options = preg_split( "/\r\n/", $_mw_adminimize_own_post_options );
							foreach ( (array) $_mw_adminimize_own_post_options as $key => $_mw_adminimize_own_post_option ) {
								$_mw_adminimize_own_post_option = trim($_mw_adminimize_own_post_option);
								array_push($metaboxes_names, $_mw_adminimize_own_post_option);
							}
							
							// pages
							$metaboxes_page = array(
								'#contextual-help-link-wrap',
								'#screen-options-link-wrap',
								'#titlediv',
								'#pageslugdiv',
								'#pagepostcustom, #pagecustomdiv, #postcustom',
								'#pagecommentstatusdiv, #commentsdiv',
								'#pagepassworddiv',
								'#pageparentdiv',
								'#pagetemplatediv',
								'#pageorderdiv',
								'#pageauthordiv, #authordiv',
								'#revisionsdiv',
								'.side-info',
								'#notice',
								'#post-body h2',
								'#media-buttons',
								'#wp-word-count',
								'#slugdiv,#edit-slug-box',
								'#misc-publishing-actions',
								'#commentstatusdiv',
								'#editor-toolbar #edButtonHTML, #quicktags'
							);

							if ( function_exists('current_theme_supports') && current_theme_supports( 'post-thumbnails', 'page' ) )
								array_push($metaboxes_page, '#postimagediv' );
							if (class_exists('SimpleTagsAdmin'))
								array_push($metaboxes_page, '#suggestedtags');
							if (class_exists('HTMLSpecialCharactersHelper'))
								array_push($metaboxes_page, '#htmlspecialchars');
							if (class_exists('All_in_One_SEO_Pack'))
								array_push($metaboxes_page, '#postaiosp, #aiosp');
							if (function_exists('tdomf_edit_post_panel_admin_head'))
								array_push($metaboxes_page, '#tdomf');
							if (function_exists('post_notification_form'))
								array_push($metaboxes_page, '#post_notification');

							$metaboxes_names_page = array(
								__('Help'),
								__('Screen Options'),
								__('Title', FB_ADMINIMIZE_TEXTDOMAIN),
								__('Permalink', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Custom Fields'),
								__('Comments &amp; Pings', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Password Protect This Page', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Attributes'),
								__('Page Template', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Page Order', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Page Author'),
								__('Page Revisions'),
								__('Related', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Messages', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('h2: Advanced Options', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Media Buttons (all)', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Word count', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Page Slug'),
								__('Publish Actions', FB_ADMINIMIZE_TEXTDOMAIN ),
								__('Discussion'),
								__('HTML Editor Button')
							);

							if ( function_exists('current_theme_supports') && current_theme_supports( 'post-thumbnails', 'page' ) )
								array_push($metaboxes_names_page, __('Page Image') );
							if (class_exists('SimpleTagsAdmin'))
								array_push($metaboxes_names_page, __('Suggested tags from', FB_ADMINIMIZE_TEXTDOMAIN ));
							if (class_exists('HTMLSpecialCharactersHelper'))
								array_push($metaboxes_names_page, __('HTML Special Characters'));
							if (class_exists('All_in_One_SEO_Pack'))
								array_push($metaboxes_names_page, 'All in One SEO Pack');
							if (function_exists('tdomf_edit_post_panel_admin_head'))
								array_push($metaboxes_names_page, 'TDOMF');
							if (function_exists('post_notification_form'))
								array_push($metaboxes_names_page, 'Post Notification');
							
							// add own page options
							$_mw_adminimize_own_page_values = _mw_adminimize_getOptionValue('_mw_adminimize_own_page_values');
							$_mw_adminimize_own_page_values = preg_split( "/\r\n/", $_mw_adminimize_own_page_values );
							foreach ( (array) $_mw_adminimize_own_page_values as $key => $_mw_adminimize_own_page_value ) {
								$_mw_adminimize_own_page_value = trim($_mw_adminimize_own_page_value);
								array_push($metaboxes_page, $_mw_adminimize_own_page_value);
							}
							
							$_mw_adminimize_own_page_options = _mw_adminimize_getOptionValue('_mw_adminimize_own_page_options');
							$_mw_adminimize_own_page_options = preg_split( "/\r\n/", $_mw_adminimize_own_page_options );
							foreach ( (array) $_mw_adminimize_own_page_options as $key => $_mw_adminimize_own_page_option ) {
								$_mw_adminimize_own_page_option = trim($_mw_adminimize_own_page_option);
								array_push($metaboxes_names_page, $_mw_adminimize_own_page_option);
							}
							
							// print menu, submenu
							if ( isset($wp_menu) && '' != $wp_menu ) {

								$i = 0;
								$x = 0;
								$class = '';
								
								$users = array( 0 => 'Profile', 1 => 'edit_users', 2 => 'profile.php', 3 => '', 4 => 'menu-top', 5 => 'menu-users', 6 => 'div' );
								//array_push( $menu, $users );
								
								foreach ($wp_menu as $item) {
									
									// non checked items
									if ( $item[2] === 'options-general.php' ) {
										//$disabled_item_adm = ' disabled="disabled"';
										$disabled_item_adm_hint = '<abbr title="' . __( 'After activate the check box it heavy attitudes will change.', FB_ADMINIMIZE_TEXTDOMAIN ) . '" style="cursor:pointer;"> ! </acronym>';
									} else {
										$disabled_item_adm = '';
										$disabled_item_adm_hint = '';
									}
									
									if ( $item[0] != '' ) {
										foreach($user_roles as $role) {
											// checkbox checked
												if ( isset( $disabled_menu_[$role]) && in_array($item[2],  $disabled_menu_[$role]) ) {
												$checked_user_role_[$role] = ' checked="checked"';
											} else {
												$checked_user_role_[$role] = '';
											}
										}
	
										echo '<tr class="form-invalid">' . "\n";
										echo "\t" . '<th>' . $item[0] . ' <span style="color:#ccc; font-weight: 400;">(' . $item[2] . ')</span> </th>';
										foreach ($user_roles as $role) {
											if ( $role != 'administrator' ) { // only admin disable items
												$disabled_item_adm = '';
												$disabled_item_adm_hint = '';
											}
											echo "\t" . '<td class="num">' . $disabled_item_adm_hint . '<input id="check_menu'. $role . $x .'" type="checkbox"' . $disabled_item_adm . $checked_user_role_[$role] . ' name="mw_adminimize_disabled_menu_'. $role .'_items[]" value="' . $item[2] . '" />' . $disabled_item_adm_hint . '</td>' . "\n";
										}
										echo '</tr>';
										
										// only for user smaller administrator, change user-Profile-File
										if ( 'users.php' === $item[2] ) {
											$x++;
											echo '<tr class="form-invalid">' . "\n";
											echo "\t" . '<th>' . __('Profile') . ' <span style="color:#ccc; font-weight: 400;">(profile.php)</span> </th>';
											foreach ($user_roles as $role) {
												echo "\t" . '<td class="num"><input disabled="disabled" id="check_menu'. $role . $x .'" type="checkbox"' . $checked_user_role_[$role] . ' name="mw_adminimize_disabled_menu_'. $role .'_items[]" value="profile.php" /></td>' . "\n";
											}
											echo '</tr>';
										}

										$x++;

										if ( !isset($wp_submenu[$item[2]]) )
											continue;

										// submenu items
										foreach ( $wp_submenu[ $item[2] ] as $subitem ) {
											$class = ( ' class="alternate"' == $class ) ? '' : ' class="alternate"';
											if ( $subitem[2] === 'adminimize/adminimize.php' ) {
												//$disabled_subitem_adm = ' disabled="disabled"';
												$disabled_subitem_adm_hint = '<abbr title="' . __( 'After activate the check box it heavy attitudes will change.', FB_ADMINIMIZE_TEXTDOMAIN ) . '" style="cursor:pointer;"> ! </acronym>';
											} else {
												$disabled_subitem_adm = '';
												$disabled_subitem_adm_hint = '';
											}
											
											echo '<tr' . $class . '>' . "\n";
											foreach ($user_roles as $role) {
												if ( isset($disabled_submenu_[$role]) )
													$checked_user_role_[$role]  = ( in_array($subitem[2], $disabled_submenu_[$role] ) ) ? ' checked="checked"' : '';
											}
											echo '<td> &mdash; ' . $subitem[0] . ' <span style="color:#ccc; font-weight: 400;">(' . $subitem[2] . ')</span> </td>' . "\n";
											foreach ($user_roles as $role) {
												if ( $role != 'administrator' ) { // only admin disable items
													$disabled_subitem_adm = '';
													$disabled_subitem_adm_hint = '';
												}
												echo '<td class="num">' . $disabled_subitem_adm_hint . '<input id="check_menu'. $role.$x .'" type="checkbox"' . $disabled_subitem_adm . $checked_user_role_[$role] . ' name="mw_adminimize_disabled_submenu_'. $role .'_items[]" value="' . $subitem[2] . '" />' . $disabled_subitem_adm_hint . '</td>' . "\n";
											}
											echo '</tr>' . "\n";
											$x++;
										}
										$i++;
										$x++;
									}
								}

							} else {
								$myErrors = new _mw_adminimize_message_class();
								$myErrors = '<tr><td style="color: red;">' . $myErrors->get_error('_mw_adminimize_get_option') . '</td></tr>';
								echo $myErrors;
							} ?>
						</tbody>
					</table>
					
					<p id="submitbutton">
						<input class="button button-primary" type="submit" name="_mw_adminimize_save" value="<?php _e('Update Options', FB_ADMINIMIZE_TEXTDOMAIN ); ?> &raquo;" /><input type="hidden" name="page_options" value="'dofollow_timeout'" />
					</p>
					<p><a class="alignright button" href="javascript:void(0);" onclick="window.scrollTo(0,0);" style="margin:3px 0 0 30px;"><?php _e('scroll to top', FB_ADMINIMIZE_TEXTDOMAIN); ?></a><br class="clear" /></p>

				</div>
			</div>
		</div>
		