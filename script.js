/**
 * Javascript to add media upload to wp4office-gf2excel settings page for forms
 * wp4office-gf2excel Gravity Forms AddOn saves form data into a given Excel
 * document and attaches it to notification emails
 * author: Dieter Pfenning (dieter.pfenning@winball.de)
 * version: 1.0
 * date: 2015-06-21
 */

jQuery(document).ready(function($){

	// add upload button to form
	jQuery('#gaddon-setting-row-excel_template_path td').append('<input type="button" name="wp4o-upload-btn" id="wp4o-upload-btn" class="button-secondary" value="' + objectL10n.submit_text + '">');
	
	
	
	// handle media upload
	jQuery('#wp4o-upload-btn').click(function(e) {
		e.preventDefault();
		var image = wp.media({ 
			title: objectL10n.submit_text,
			multiple: false
		}).open()
		.on('select', function(e){
			// This will return the selected image from the Media Uploader, the result is an object
			var uploaded_image = image.state().get('selection').first();
			// We convert uploaded_image to a JSON object to make accessing it easier
			var image_url = uploaded_image.toJSON().url;
			// remove protocol and domain from image_url
			image_url = image_url.replace(/^.*\/\/[^\/]+/, '');
			// Let's assign the url value to the input field
			jQuery('#excel_template_path').val(image_url);
		});
	});
	
	
});