jQuery(document).ready( function(){
	var qs
	jQuery('#afw').on('click','.executeMacro', function(){
		//alert('execute '+jQuery(this).attr('id'));
		qs = 'srcFormId=110&button=execute&macroId='+jQuery(this).attr('macroId');
		afwUpdateFrame( 'themeManagerAjaxHtmlHandler', qs, 'themeManagerMacros' );
		//cj.remote({
		//	'method':'themeManagerAjaxHtmlHandler'
		//	,'destinationId': 'themeManagerMacros'
		//	,'queryString': qs
		//});
	});
})
