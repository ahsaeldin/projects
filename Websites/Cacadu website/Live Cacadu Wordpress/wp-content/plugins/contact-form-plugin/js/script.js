jQuery(document).ready(function(){
	jQuery('#cntctfrm_additions_options').change( function() {
		if(jQuery(this).is(':checked') )
			jQuery('.cntctfrm_additions_block').removeClass('cntctfrm_hidden');
		else
			jQuery('.cntctfrm_additions_block').addClass('cntctfrm_hidden');
	});
	jQuery('#cntctfrm_change_label').change( function() {
		if(jQuery(this).is(':checked') )
			jQuery('.cntctfrm_change_label_block').removeClass('cntctfrm_hidden');
		else
			jQuery('.cntctfrm_change_label_block').addClass('cntctfrm_hidden');
	});
	jQuery('#cntctfrm_display_add_info').change( function() {
		if(jQuery(this).is(':checked') )
			jQuery('.cntctfrm_display_add_info_block').removeClass('cntctfrm_hidden');
		else
			jQuery('.cntctfrm_display_add_info_block').addClass('cntctfrm_hidden');
	});
});