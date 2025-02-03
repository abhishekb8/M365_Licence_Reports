A simple powershell script uses Graph API, to fetch licences to the specified tenant.
The script uses MSgraph to fetch the report.

There are 3 parts to the script.

Part 1 :  Authentication
          For autntication it uses Azure App with graph read only permissions

Part 2 : Report
         The graph queries M365 tenant for the entitled and other applicable licecnes in a list

Part 3 : Email
         Converts the output to an HTML format and sends out an email to desired groups or email addresses.
