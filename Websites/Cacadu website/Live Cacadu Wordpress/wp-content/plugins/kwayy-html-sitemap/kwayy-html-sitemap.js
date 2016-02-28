jQuery(document).ready(function(){
	jQuery( "#kwayyhs-sortable" ).sortable({
		placeholder: "kwayyhs-ui-state-highlight",

		update: function(event, ui) {
				var fruitOrder = jQuery("#kwayyhs-sortable").sortable('toArray').toString();
				jQuery('#kwayyhs-sortorder').val(fruitOrder);
				//$.get('update-sort.cfm', {fruitOrder:fruitOrder});
			}
	});
	jQuery( "#kwayyhs-sortable" ).disableSelection();
});
