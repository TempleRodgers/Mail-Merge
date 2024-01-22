# Mail-Merge
Simple Mail merge using Google Docs template and Google Sheets list. This takes a form letter and
merges it with data from a spreadsheet to create either a single document with multiple custom 
copies of the template or multiple documents that are separate custom copies of the template.

This needs a google docs template that has fields e.g.:
--------------------------------------------------------------------------------------------------------
                                                                    
                                                                    {{SENDER_NAME}}
                                                                    {{SENDER_ADDRESS}}
                                                                    {{SENDER_POSTCODE}}
                                                                    {{SENDER_PHONE}}
                                                                    {{SENDER_EMAIL}}
                                                                    {{DATE}}
{{TO_NAME}}
{{TO_TITLE}}
{{TO_COMPANY}}
{{TO_ADDRESS}}
Dear {{TO_NAME}},
Type the body text here, you can embed fields such as {{TO_NAME}} and {{SENDER_NAME}} in the body text.
Make a copy of the spreadsheet then, in the copy (you can rename) amend the sender and sender_details tab and mail_merge tab. You can change headings as you see fit, the script will work this out.
Sincerely,



{{SENDER_NAME}}
--------------------------------------------------------------------------------------------------------
Once you've created the template, go to Tools -> Apps Script and copy/paste the two files: the Apps Script
.gs file and the HTML .html file.

You will need to close and open the doc for it to run the initial setup of the script and ask you for 
permission to run the code. The code prompts you to select a spreadsheet, which you can select from 
adjacent google sheets files in google drive.

The code does not do anything it shouldn't, there is no adware.

You will need a google sheet with two tabs: Mail_Merge where all the field data is kept and 
Sender_Details where all the sender's information is kept. The column titles must be the 
correct case and the tabs must match Mail_Merge and Sender_Details.

The fields are completely flexible - call them what you want and have as many columns as you like.

