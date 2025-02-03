A simple powershell script uses Graph API, to fetch licences to the specified tenant.
The script uses MSgraph to fetch the report.

There are 3 parts to the script.

Part 1 :  Authentication, uses Azure App with graph read only permissions

Part 2 : For report queries Graph API for the M365 tenant for the entitled and other applicable licecnes in a list and converts SKUId's to Friendly Names

Part 3 : Sends an Email, Converts the output to an HTML format and sends out an email to desired groups or email addresses
