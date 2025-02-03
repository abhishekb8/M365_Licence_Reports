#Import required Modules
Import-Module -Name Microsoft.Graph    

# Define the Application (Client) ID and Secret
$ApplicationClientId = '<Azureapp_ID>'
$ApplicationClientSecret = '<Azureapp_Secret>'
$TenantId = '<Tenant_ID>'

# Convert the Client Secret to a Secure String
$SecureClientSecret = ConvertTo-SecureString -String $ApplicationClientSecret -AsPlainText -Force

# Create a PSCredential Object Using the Client ID and Secure Client Secret
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationClientId, $SecureClientSecret

# Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

# Step 1: Define SkuPartNumber to Friendly Name Mapping
$SkuFriendlyNames = @{

"AAD_PREMIUM" = "Microsoft Entra ID P1"
"CRM_HYBRIDCONNECTOR" = "Dynamics 365 Hybrid Connector"
"DEFENDER_ENDPOINT_P1" = "Microsoft Defender for Endpoint P1"
"EMS" = "Enterprise Mobility + Security E3"
"ENTERPRISEPACK" = "Office 365 E3"
"EXCHANGEENTERPRISE" = "Exchange Online (Plan 2)"
"EXCHANGESTANDARD" = "Exchange Online (Plan 1)"
"FLOW_FREE" = "Microsoft Power Automate Free"
"GUIDES_USER" = "Dynamics 365 Guides"
"IDENTITY_THREAT_PROTECTION" = "Microsoft 365 E5 Security"
"INTUNE_A_D" = "Microsoft Intune Device"
"INTUNE_A_VL" = "Microsoft Intune Plan 1 A"
"MCOCAP" = "Microsoft Teams Shared Devices"
"MCOEV" = "Microsoft Teams Phone Standard"
"MCOMEETADV" = "Microsoft 365 Audio Conferencing"
"MEETING_ROOM" = "Microsoft Teams Rooms Standard"
"Microsoft_365_Copilot" = "Microsoft Copilot for Microsoft 365"
"Microsoft_365_E3_Extra_Features" = "Microsoft 365 E3 Extra Features"
"MICROSOFT_REMOTE_ASSIST" = "Dynamics 365 Remote Assist"
"Microsoft_Teams_Audio_Conferencing_select_dial_out" = "Microsoft Teams Audio Conferencing with dial-out to USA/CAN"
"Microsoft_Teams_Premium" = "Microsoft Teams Premium Introductory Pricing"
"Microsoft_Teams_Rooms_Pro" = "Microsoft Teams Rooms Pro"
"OFFICE365_MULTIGEO" = "Multi-Geo Capabilities in Office 365"
"PBI_PREMIUM_P1_ADDON" = "Power BI Premium P1"
"PBI_PREMIUM_P2_ADDON" = "Power BI Premium P2"
"PBI_PREMIUM_PER_USER" = "Power BI Premium Per User"
"PHONESYSTEM_VIRTUALUSER" = "Microsoft Teams Phone Resource Account"
"POWER_BI_PRO" = "Power BI Pro"
"POWER_BI_PRO_DEPT" = "Power BI Pro Dept"
"POWER_BI_STANDARD" = "Microsoft Fabric (Free)"
"POWERAPPS_DEV" = "Microsoft Power Apps for Developer"
"POWERAPPS_INDIVIDUAL_USER" = "Power Apps and Logic Flows"
"POWERAPPS_VIRAL" = "Microsoft Power Apps Plan 2 Trial"
"POWERAUTOMATE_ATTENDED_RPA" = "Power Automate Premium"
"POWERAUTOMATE_UNATTENDED_RPA" = "Power Automate unattended RPA add-on"
"PROJECT_PLAN1_DEPT" = "Project Plan 1 (for Department)"
"PROJECT_PLAN3_DEPT" = "Project Plan 3 (for Department)"
"PROJECTPREMIUM" = "Project Online Premium"
"PROJECTPROFESSIONAL" = "Project Plan 3"
"RIGHTSMANAGEMENT_ADHOC" = "Rights Management Adhoc"
"SHAREPOINTENTERPRISE" = "SharePoint Online (Plan 2)"
"SHAREPOINTSTANDARD" = "SharePoint Online (Plan 1)"
"SPE_E3" = "Microsoft 365 E3"
"SPE_E5" = "Microsoft 365 E5"
"SPE_F1" = "Microsoft 365 F3"
"STREAM" = "Microsoft Stream"
"Teams_Premium_(for_Departments)" = "Teams Premium (for Departments)"
"VISIO_PLAN2_DEPT" = "Visio Plan 2"
"VISIOCLIENT" = "Visio Online Plan 2"
"VISIOONLINE_PLAN1" = "Visio Online Plan 1"
"Viva_Goals_User_led" = "Viva Goals User-led"
"WACONEDRIVESTANDARD" = "OneDrive for Business (Plan 1)"
"Win10_VDA_E3" = "Windows 10/11 Enterprise E3"
"WINDOWS_STORE" = "Windows Store for Business"
"PROJECT_P1" = "Project Plan 1"
"Cross_tenant_user_data_migration" = "Cross Tenant Data Migration"
}

# Step 2: Retrieve License Information
$LicenseData = Get-MgSubscribedSku

# Step 3: Convert SkuPartNumber to Friendly Names and Format Output
$FormattedLicenseData = $LicenseData | Select-Object @{Name="FriendlyName";Expression={$SkuFriendlyNames[$_.SkuPartNumber]}}, @{Name="Total Licenses";Expression={$_.PrepaidUnits.Enabled}}, ConsumedUnits,@{Name="Available Licenses";Expression={$_.PrepaidUnits.Enabled - $_.ConsumedUnits }}

# Step 4: Display the Output
$FormattedLicenseData

# Step 5: Build the HTML body
$commandOutput = $FormattedLicenseData | ConvertTo-Html -Fragment
$date = Get-Date -Format dd-MMM-yyyy
$htmlBody = @"
<html>
<head>
    <style>
        body {font-family: Arial, sans-serif;}
        table {width: 100%; border-collapse: collapse;}
        th, td {padding: 8px; text-align: left; border: 1px solid #ddd;}
        th {background-color: #f2f2f2;}
    </style>
</head>
<body>
    <h2>M365 Licence Summary - $date</h2>
    $commandOutput
</body>
</html>
"@


# Step 6: Define the email parameters
$Server = "<SMTP_Server>"
$From = 'M365LicenceStatus@emaildomain.com'
$To = 'reciever@emaildomain.com'
$Bcc = 'recieverbcc@maildomain.com'
$Subject = "M365 License Usage Summary"
$Body = $htmlBody

# Step 7: Send the email
Send-MailMessage -SmtpServer $Server -Port 25 -From $From -To $To -Bcc $Bcc -Subject $Subject -Body $Body -BodyAsHtml

# Step 8: Disconnect Graph
Disconnect-MgGraph
