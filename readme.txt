=== Gravity Forms To Excel AddOn ===
Contributors: 
Tags: gravityforms, excel, excel export, forms, attachment, email, notification, no CSV
Donate link: 
Requires at least: 3.7
Tested up to: 4.2.2
Stable tag: 0.1.3
License: GPLv2 or later
License URI: http://www.gnu.org/licenses/gpl-2.0.html

Gravity Forms AddOn which saves form data into a given Excel document and attaches it to notification emails



== Description ==
This Gravity Forms AddOn saves form data into a given Excel document and attaches it to notification emails. You don't need any programming skills to get native Excel documents back as the result of your Gravity Forms web form. After uploading your Excel 2007 file (.xslx, other versions are not supported) the form data is saved into one sheet (which you can define) of your document. You can then select to which notification emails this Excel file should be attached to. Using simple Excel formulas (=A1) you can fill out complex Excel sheets with data from the web form. No further export or import of CSV data is required.

= Operating instructions =

1. Create your form with Gravity Forms
1. Give all your fields admin field labels (under the tab "Advanced")
1. Create your notification emails
1. Open the GF2Excel form settings, upload your Excel file, type in your sheet number to insert the form data and finally select the notifications you would like to attach the filled in Excel file.
1. Submit your form and open your Excel file. Connect your actual form fields through formulas (=A1) with data of the sheet which is filled by Gravity Forms. The admin field labels will help you to associate the data with the form fields.
1. Open the WP4O-GF2Excel form settings again and upload the Excel file with your formulas.
1. Repeat steps 5 and 6 until you are satisfied with the result.
1. Be happy :-)


*This plugin requires Gravity Forms by RocketGenius to be active.*

*This plugin was successfully tested on WordPress Multisite without any abnormalities.*

This plugin is maintained [by winball.de](http://winball.de) on [wp4office.winball2.de](http://wp4office.winball2.de/gf2excel). We welcome your pull requests, comments and suggestions for improvement. Additional [help and example files](http://wp4office.winball2.de/gf2excel/help) with descriptions are available. You can [try out a demo](http://wp4office.winball2.de/gf2excel/demo) before installing.

= You do have problems or need individual service? =
Professional web services are our actual business. If you need help with your form or your Excel file, please feel free to [contact us](http://winball.de/wp4office-gf2excel-services).



== Installation ==
Upload the plugin to your blog (manually via ftp or through the dashboard) and activate it.

*This plugin requires Gravity Forms by RocketGenius to be active.*



== Frequently Asked Questions ==
We are waiting for input ...



== Screenshots ==
1. Use the admin field labels to associate the data with the form fields.
2. Upload your Excel file, type in your sheet number and select the notifications.
3. The target sheet will contain the admin field labels in column A and the form data in column B.
4. Using simple formulas (e.g. A1=...) your form will show the data coming from the web form.



== Changelog ==
= 0.1.3 = 
* make first sheet (index 0) active, not the sheet, where the form data is stored
* columns A and C switched, to allow the use of VLOOKUP in Excel
* validate user input and create error messages when required
* language support added (I18n)
* Gift Certificate sample added

= 0.1.2 = 
* fixed issue with no selected notifications in wp4o-gf2excel and blank screen after sending form

= 0.1.1 =
* Initial release.