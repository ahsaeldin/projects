<?php 
/*

function remove_menus () {

global $menu;
	$restricted = array(__('Posts'), __('Media'), __('Links'), __('Appearance'), __('Tools'), __('Users'), __('Settings'), __('Comments'), __('Plugins'));
	end ($menu);
	while (prev($menu)){
		$value = explode(' ',$menu[key($menu)][0]);
		if(in_array($value[0] != NULL?$value[0]:"" , $restricted)){unset($menu[key($menu)]);}
	}
}
add_action('admin_menu', 'remove_menus');


function remove_submenus() {
global $submenu;
unset($submenu['index.php'][10]); // Removes 'Updates'.
//unset($submenu['themes.php'][5]); // Removes 'Themes'.
//unset($submenu['options-general.php'][15]); // Removes 'Writing'.
//unset($submenu['options-general.php'][25]); // Removes 'Discussion'.
//unset($submenu['edit.php'][11]); // Removes 'Tags'.


global $submenu;
unset($submenu[10]);

}

add_action('admin_menu', 'remove_submenus');
*/
 ?>