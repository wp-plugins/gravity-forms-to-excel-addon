<?php
/*
Plugin Name: Gravity Forms To Excel AddOn
Plugin URI: http://wp4office.winball2.de/gf2excel
Description: Gravity Forms AddOn which saves form data into a given Excel document and attaches it to notification emails
Version: 0.1.3
Author: winball.de
Author URI: http://winball.de/
License: GPLv2 or later
Text Domain: gf2excel-addon

------------------------------------------------------------------------
Copyright 2012-2015 winball.de (PG Consulting GmbH)

Publishing date: 2015-06-14 22:47:52

This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 2 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
*/



//------------------------------------------
if (!class_exists("GFForms")) {
	die();
} else {
	GFForms::include_addon_framework();

	class WP4O_GF2Excel extends GFAddOn {

		protected $_version = "1.1";
		protected $_min_gravityforms_version = "1.7.9999";
		protected $_slug = "gf2excel-addon";
		protected $_path = "gf2excel-addon/gf2excel-addon.php";
		protected $_full_path = __FILE__;
		protected $_title = "GravityForms2Excel";
		protected $_short_title = "GF2Excel";
		
		
		
		public function init(){
			parent::init();
			
			// ********** add excel (xlsx) mime type, if not already allowed **********
			$wp4o_allowed_mime_types = get_allowed_mime_types();
			if (!array_key_exists('xlsx', $wp4o_allowed_mime_types)) {
				function wp4o_add_excel_mime_type( $existing_mimes ) {
					$existing_mimes['xlsx'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
					return $existing_mimes;
				}
				add_filter('upload_mimes', 'wp4o_add_excel_mime_type');
			}
			
			
			
			// ********** add media upload to plugin forms settings page **********
			// enqueue javascripts and styles (required for media upload)
			function wp4o_enqueue(){
				if( $_GET['page'] == 'gf_edit_forms' && $_GET['subview'] == 'gf2excel-addon' ){
					wp_enqueue_script('jquery');
					wp_enqueue_media();
					wp_enqueue_script( 'gf2excel-addon',plugin_dir_url( __FILE__ ).'script.js', array( 'gform_gravityforms' ), '0.1.1', true );
					// add localization to javascript
					wp_localize_script( 'gf2excel-addon', 'objectL10n', array(
						'submit_text' => __( 'Upload Excel file', 'gf2excel-addon' ),
					) );
					//wp_enqueue_style( 'wp4office-gf2excel', plugin_dir_url( __FILE__ ).'style.css', array(), '0.1.1' );
				}
			}
			add_action( 'admin_enqueue_scripts', 'wp4o_enqueue');
			
			
			
			// ********** add temp excel as attachment to notification emails **********
			function wp4o_attach_excel($notification, $form, $entry){
				// read plugin settings for this form
				$settings = $form['gf2excel-addon'];
				
				// read active notifications for this form
				$dips_notifications = array();
				foreach ($form['notifications'] as $key => $value){
					if ($value['isActive']==1 && $settings[$key]==1){
						$dips_notifications[] = $key;
					}
				}
				
				// check if we have something to do for this form or we may exit
				if (empty($dips_notifications) || !in_array($notification['id'], $dips_notifications))
					return $notification;
				
				// extract labels/adminLabel from form array to find id/adminlabel much faster and easier
				$dips_form_fields = $form['fields'];//array
				$dips_lables = array();
				foreach ($dips_form_fields as $value) {
					$dips_lables[$value['id']] = $value['adminLabel'];
				}
				
				// include phpExcel
				require_once(__DIR__.'/includes/PHPExcel.php');
				// include file phpexcel I/O
				require_once (__DIR__.'/includes/PHPExcel/IOFactory.php');
				
				// open excel template file into excel object and set active sheet by index
				$dips_excel = PHPExcel_IOFactory::createReader('Excel2007');
				try {
					$dips_excel = $dips_excel->load(ABSPATH.$settings['excel_template_path']);
					$old_sheet_index = $dips_excel->getActiveSheetIndex();
				}
				catch (Exception $e) {
					return $notification; //file does not exist, so we exit
				}
				try {
					$dips_excel->setActiveSheetIndex(intval($settings['excel_sheet_index']));
				}
				catch (Exception $e) {
					return $notification; //sheet does not exist, so we exit
				}
				
				// loop all entries and save into excel object
				$i = 1;
				foreach ($entry as $key => $value) {	
					$dips_excel->getActiveSheet()->setCellValue('C'.$i, $key) ->setCellValue('B'.$i, $value) ->setCellValue('A'.$i, $dips_lables[$key]);
					$i++;
				}
				
				// write data back into temp excel file
				$dips_excel->setActiveSheetIndex($old_sheet_index); // set to first sheet
				$objWriter = PHPExcel_IOFactory::createWriter($dips_excel, 'Excel2007');
				$excel_file_name = pathinfo ( $settings['excel_template_path'] , PATHINFO_FILENAME );// no .xlsx
				$filename = $excel_file_name.'_wp4o_'.$entry['id'].'.xlsx';// add entry id to end of filename
				try {
					$objWriter->save(__DIR__.'/tmp/'.$filename);
					// try to close file and object
					$dips_excel->disconnectWorksheets();
					unset($objWriter, $dips_excel);
				}
				catch (Exception $e) {
					return $notification; //file cannot be saved, so we exit
				}
				
				
				// attach previous saved excel file to notification email
				$notification['attachments'] = array(__DIR__.'/tmp/'.$filename);
				return $notification;
			}
			add_filter( 'gform_notification', 'wp4o_attach_excel', 10, 3 );
			
			
			
			// ********** delete temp excel file after sending notifications **********
			function wp4o_delete_excel($entry, $form){
				// read plugin settings for this form
				$settings = $form['gf2excel-addon'];
				
				// read active notifications for this form
				$dips_notifications = array();
				foreach ($form['notifications'] as $key => $value){
					if ($value['isActive']==1 && $settings[$key]==1){
						$dips_notifications[] = $key;
					}
				}
				
				// check if we have something to do for this form or we may exit
				if (empty($dips_notifications))
					exit();
				
				// delete file
				$excel_file_name = pathinfo ( $settings['excel_template_path'] , PATHINFO_FILENAME );// no extension, no .xlsx
				$filename = $excel_file_name.'_wp4o_'.$entry['id'].'.xlsx';
				
				try {
					unlink( __DIR__.'/tmp/'.$filename );
				}
				catch (Exception $e) {
					exit(); //file does not exist or cannot be deleted
				}
			}
			add_action( 'gform_after_submission', 'wp4o_delete_excel', 10, 2 );
		}
		
		
		
		public function plugin_page() {
			_e("<p>Thank you for using this plugin. If you like it, we would like to invite you to rate it on <a href='https://wordpress.org/plugins/wp4office-gf2excel/' target='_blank'>wordpress.org</a>.</p><h2>Description</h2> <p>This Gravity Forms AddOn saves form data into a given Excel document and attaches it to notification emails.You don't need any programming skills to get native Excel documents back as the result of your Gravity Forms web form. After uploading your Excel 2007 file (.xslx, other versions are not supported) the form data is saved into one sheet (which you can define) of your document. You can then select to which notification emails thisExcel file should be attached to. Using simple Excel formulas (=A1)you can fill out complex Excel sheets with data from the web form.No further export or import of CSV data is required.</p><h2>Operating instructions</h2> <ol> <li>Create your form withGravity Forms</li> <li>Give all your fields admin field labels(under the tab 'Advanced')</li> <li>Create your notification emails</li> <li>Open the WP4O-GF2Excel form settings, upload yourExcel file, type in your sheet number to insert the form data and finally select the notifications you would like to attach the filled in Excel file.</li> <li>Submit your form and open your Excel file.Connect your actual form fields through formulas (=A1) with data of the sheet which is filled by Gravity Forms. The admin field labels will help you to associate the data with the form fields.</li><li>Open the WP4O-GF2Excel form settings again and upload the Excel file with your formulas.</li> <li>Repeat steps 5 and 6 until you are satisfied with the result.</li> <li>Be happy :-)</li> </ol><p><i>This plugin requires Gravity Forms by RocketGenius to be active.</i></p><p><i>This plugin was successfully tested on WordPress Multisite without any abnormalities.</i></p><p>This plugin is maintained by <a href='http://winball.de' target='_blank'>winball.de</a> on <a href='http://wp4office.winball2.de/gf2excel' target='_blank'>wp4office.winball2.de</a>. We welcome your pull requests,comments and suggestions for improvement. Additional <a href='http://wp4office.winball2.de/gf2excel/help' target='_blank'>help and example files</a> with descriptions are available. You can <a href='http://wp4office.winball2.de/gf2excel/demo' target='_blank'>try out a demo</a> before installing. Technical support is available under <a href='https://wordpress.org/support/plugin/wp4office-gf2excel' target='_blank'>wordpress.org.</a></p><h2>You do have problems or need individual service?</h2><p>Professional web services are our actual business. If you need help with your form or your Excel file, please feel free to <a href='http://winball.de/wp4office-gf2excel-services' target='_blank'>contact us</a>.</p>",'gf2excel-addon');
		}
		
		
		
		public function form_settings_fields($form) {
			// get active notifications of the form
			$dips_choices = array();
			foreach ($form['notifications'] as $key => $value){
				if ($value['isActive']==1){
					$dips_choices[] = array("label" => $value['name'],"name"  => $key);
				}
			}
			
			return array(
				array(
					"title"  => "GF2Excel-Addon Settings",
					"fields" => array(
						array(
							"label"               => __("Excel Template file:",'gf2excel-addon'),
							"type"                => "text",
							"name"                => "excel_template_path",
							"tooltip"             => __("Insert the path to the excel template file (must be Excel 2007 and end with .xlsx; should start with /wp-content/...)",'gf2excel-addon'),
							"class"               => "medium",
							"validation_callback" => array($this, "dips_validate_excel_template_path")//feedback_callback does not fire "there was an error.."
						),
						array(
							"label"   => __("Excel sheet index:",'gf2excel-addon'),
							"type"    => "text",
							"name"    => "excel_sheet_index",
							"tooltip" => __("Insert the sheet index of the excel sheet where you want to save the from data (excel sheet indices start with 0, please use indices which do exist)",'gf2excel-addon'),
							"class"   => "medium",
							"validation_callback" => array($this, "dips_validate_excel_sheet_index")
						),
						array(
							"label"   => __("Add Excel file as attachment to notification emails:",'gf2excel-addon'),
							"type"	=> "checkbox",
							"name"	=> "excel_notifications",
							"tooltip" => __("Select the notification emails, where you want to attach the excel file",'gf2excel-addon'),
							"choices" => $dips_choices,
						)
					)
				)
			);
		}
		
		
		public function dips_validate_excel_template_path($value){
			$excel_template_path = GFAddOn::get_setting('excel_template_path');
			$excel_notifications = array_slice(GFAddOn::get_current_settings(), 2);// after the two first settings follow the notification ids
			// check for .xlsx extension
			if ($excel_template_path!='' && pathinfo($excel_template_path , PATHINFO_EXTENSION)!='xlsx'){
				GFAddOn::set_field_error( array('name' =>'excel_template_path'), __('.xlsx extension is missing','gf2excel-addon') );
				return false; // .xlsx extension missing
			}
			
			// check for Excel 2007
			require_once(__DIR__.'/includes/PHPExcel.php');
			require_once (__DIR__.'/includes/PHPExcel/IOFactory.php');
			try {
				$filetype = PHPExcel_IOFactory::identify(ABSPATH.$excel_template_path);
			}
			catch (Exception $e) {
				$filetype='file not found'; //file does not exist
			}
			if ( ($excel_template_path!='' && !$filetype=='Excel2007') ){
				GFAddOn::set_field_error( array('name' =>'excel_template_path'), __('The file type is not Excel 2007 or the file does not exist','gf2excel-addon') );
				return false; // file type is not excel 2007 or file does not exist
			}
			
			// check if at least one notification is selected and we do have an existing excel file
			if ( (array_search('1', $excel_notifications)) && ($filetype!='Excel2007') ){
				GFAddOn::set_field_error( array('name' =>'excel_template_path'), __('An Excel 2007 file must be uploaded to attach it to the selected notifications','gf2excel-addon') );
				return false;//an excel file must be uploaded to attach it to the selected notifications
			}
			
			// check if file path points to an existing file
			if ($excel_template_path!='' && !file_exists ( ABSPATH.$excel_template_path )){
				GFAddOn::set_field_error( array('name' =>'excel_template_path'), __('There is no Excel file under your specified path','gf2excel-addon') );
				return false;//There is no Excel file under your specified path
			}
			
			return true;
		}
		
		
		public function dips_validate_excel_sheet_index($value){
			$excel_template_path = GFAddOn::get_setting('excel_template_path');
			$excel_sheet_index = GFAddOn::get_setting('excel_sheet_index');
			if ($excel_sheet_index!='' && intval($excel_sheet_index)>=0 && $excel_template_path!='') {
				// check if sheet index exists
				require_once(__DIR__.'/includes/PHPExcel.php');
				require_once (__DIR__.'/includes/PHPExcel/IOFactory.php');
				$dips_excel = PHPExcel_IOFactory::createReader('Excel2007');
				try {
					$dips_excel = $dips_excel->load(ABSPATH.$excel_template_path);
					$sheetCount = $dips_excel->getSheetCount()-1;// because index starts with 0
					// try to close file and object
					$dips_excel->disconnectWorksheets();
				}
				catch (Exception $e) {
					GFAddOn::set_field_error( array('name' =>'excel_sheet_index'), __('An Excel file must exist to define sheet index','gf2excel-addon') );
					return false; // excel file must exist to define sheet index
				}

				if (intval($excel_sheet_index) > $sheetCount){
					GFAddOn::set_field_error( array('name' =>'excel_sheet_index'), __('A sheet with that index does not exist in your Excel file','gf2excel-addon') );
					return false; // sheet with that index does not exist
				}
			}
			return true;
		}
	
	
	}
	
	new WP4O_GF2Excel();
	
}