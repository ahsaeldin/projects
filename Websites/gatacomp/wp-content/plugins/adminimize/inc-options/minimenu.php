<?php
// options for write page
?>

		<?php screen_icon('tools'); ?>
		<h2><?php _e('Adminimize', FB_ADMINIMIZE_TEXTDOMAIN ); ?></h2>
		<div id="poststuff" class="metabox-holder has-right-sidebar">
			
			<div id="side-info-column" class="inner-sidebar">
				<div class="meta-box-sortables">
					<div id="about" class="postbox ">
						<div class="handlediv" title="<?php _e('Click to toggle'); ?>"><br/></div>
						<h3 class="hndle" id="about-sidebar"><?php _e('About the plugin', FB_ADMINIMIZE_TEXTDOMAIN ) ?></h3>
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
						</div>
					</div>
				</div>
			</div>
			
			<div id="post-body" class="has-sidebar">
				<div id="post-body-content" class="has-sidebar-content">
					<div id="normal-sortables" class="meta-box-sortables">
						<div id="about" class="postbox ">
							<div class="handlediv" title="<?php _e('Click to toggle'); ?>"><br/></div>
							<h3 class="hndle" id="menu"><?php _e('MiniMenu', FB_ADMINIMIZE_TEXTDOMAIN ) ?></h3>
							<div class="inside">
								<table class="widefat" cellspacing="0">
									<tr class="alternate">
										<td class="row-title"><a href="#backend_options"><?php _e('Backend Options', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr>
										<td class="row-title"><a href="#global_options"><?php _e('Global options', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr class="alternate">
										<td class="row-title"><a href="#config_menu"><?php _e('Menu Options', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr>
										<td class="row-title"><a href="#config_edit_post"><?php _e('Write options - Post', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr class="alternate">
										<td class="row-title"><a href="#config_edit_page"><?php _e('Write options - Page', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr>
										<td class="row-title"><a href="#links_options"><?php _e('Links options', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr class="alternate">
										<td class="row-title"><a href="#nav_menu_options"><?php _e('WP Nav Menu', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr>
										<td class="row-title"><a href="#set_theme"><?php _e('Set Theme', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr class="alternate">
										<td class="row-title"><a href="#import"><?php _e('Export/Import Options', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr>
										<td class="row-title"><a href="#uninstall"><?php _e('Deinstall Options', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
									<tr class="alternate">
										<td class="row-title"><a href="#about"><?php _e('About the plugin', FB_ADMINIMIZE_TEXTDOMAIN ); ?></a></td>
									</tr>
								</table>
							</div>
						</div>
					</div>
				</div>
			</div>
			<br class="clear"/>
		</div>
		