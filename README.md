![LOGO](https://github.com/davecra/outlook-sample-1/blob/master/assets/logo-filled.png?raw=true)
# Outlook-Block-Dialog-Expand-Sample
This sample is an update of a previous Outlook Sample showing only the block on send capability. This add-in does the following:
* Grabs all users on the To/CC/BCC lines
* Splits any groups found (and groups in groups)
* Determines if there are any external email addresses compared to the current users domain
* If there are external users, it prompts the user with a Yes/No/Cancel dialog to verify if they want to send.
## Requirements
This project uses:
* [easyEws](https://github.com/davecra/easyEws)
* [OfficeJS.dialogs](https://github.com/davecra/officejs.dialogs)
## Additional Information
This project is for a demo at the MVP Summit, Monday March 5th at 1:30pm.
