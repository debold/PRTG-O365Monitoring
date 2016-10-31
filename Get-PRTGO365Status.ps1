<#
.SYNOPSIS
Retrieves current service information from Office 365 tenant in PRTG compatible format

.DESCRIPTION
The Get-PRTGO365Status.ps1 uses Microsofts REST api to get the current health status of your Office 365 tenant. The XML output can be used as PRTG custom sensor.

.PARAMETER ClientID 
Represents the ClientId that is used to connect to your Office 365 tenant. See NOTES section for more details.

.PARAMETER ClientSecret
Represents the corresponding client secret to connect to your Office 365 tenant. See NOTES section for more details.

.PARAMETER TenantIdentifier
Represents the tenant to be monitored. Not the tenant name used in your Office 365 URL (e.g. https://yourtenant.onmicrosoft.com)

.EXAMPLE
Retrieves Office 365 health information for specified tenant.
Get-PRTGO365Status.ps1 -ClientId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" -ClientSecret "StrongPasswordFromAzureActiveDirectory" -TenantIdentifier "ffffffff-gggg-hhhh-iiii-jjjjjjjjjjjj"

.NOTES
Your tenant needs to be prepared to access the health information. Detailed configuration guidance can be found under http://www.team-debold.de/2016/07/22/prtg-office-365-status-ueberwachen/
For the lookup in PRTG to work you need to copy the file "custom.office365.value.ovl" to your PRTG installation folder (/lookups/custom/) of your core server and reload the lookups 
(Setup/System Administration/Administrative Tools -> Load Lookups).

Author:  Marc Debold
Version: 1.2
Version History:
    1.2  31.10.2016  Code tidy
                     Changed error handling to function
                     Changed script name
                     Changed XML output
    1.1  06.08.2016  Corrected naming mismatch in ovl file (thanks to playordie)
                     Added -UseBasicParsing Parameter to Invoke-WebRequest to bypass uninitialized Internet Explorer (thanks to playordie)
    1.0  22.07.2016  Initial release

For further reading:
    Result definition for service health: https://samlman.wordpress.com/2016/03/18/the-office365mon-rest-apis-continue-to-grow/
    PowerShell Snippets for O365 health monitoring: https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/O365InvestigationDataAcquisition.ps1
    Prerequisites for O365 monitoring: https://msdn.microsoft.com/EN-US/library/office/dn707383.aspx
    More information about prerequisites: https://azure.microsoft.com/de-de/documentation/articles/active-directory-application-objects/#BKMK_AppObject
    O365 tenant id: https://support.office.com/de-de/article/Suchen-Ihrer-Office-365-Mandanten-ID-6891b561-a52d-4ade-9f39-b492285e2c9b

.LINK
http://www.team-debold.de/2016/07/22/prtg-office-365-status-ueberwachen/

#>
[CmdletBinding()] param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $ClientID,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $ClientSecret,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $TenantIdentifier
)

$Resource = "https://manage.office.com"
$LoginURL = "https://login.windows.net"

$ReturnCodes = @{
    ServiceOperational = 0;
    Canceled = 1;
    Completed= 2;
    InProgress = 3;
    Scheduled = 4;
    ExtendedRecovery = 5;
    ServiceInterruption = 6;
    ServiceDegradation = 7;
    PostIncidentReviewPublished = 8;
    ServiceRestored = 9;
    VerifyingService = 10;
    RestoringService = 11;
    Investigating = 12
}

<# Function to raise error in PRTG style and stop script #>
function New-PrtgError {
    [CmdletBinding()] param(
        [Parameter(Position=0)][string] $ErrorText
    )

    Write-Host "<PRTG>
    <Error>1</Error>
    <Text>$ErrorText</Text>
</PRTG>"
    Exit
}

function Out-Prtg {
    [CmdletBinding()] param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)][array] $MonitoringData
    )
    # Create output for PRTG
    $XmlDocument = New-Object System.XML.XMLDocument
    $XmlRoot = $XmlDocument.CreateElement("prtg")
    $XmlDocument.appendChild($XmlRoot) | Out-Null
    # Cycle through outer array
    foreach ($Result in $MonitoringData) {
        # Create result-node
        $XmlResult = $XmlRoot.appendChild(
            $XmlDocument.CreateElement("result")
        )
        # Cycle though inner hashtable
        $Result.GetEnumerator() | ForEach-Object {
            # Use key of hashtable as XML element
            $XmlKey = $XmlDocument.CreateElement($_.key)
            $XmlKey.AppendChild(
                # Use value of hashtable as XML textnode
                $XmlDocument.CreateTextNode($_.value)    
            ) | Out-Null
            $XmlResult.AppendChild($XmlKey) | Out-Null
        }
    }
    # Format XML output
    $StringWriter = New-Object System.IO.StringWriter 
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
    $XmlWriter.Formatting = "indented"
    $XmlWriter.Indentation = 1
    $XmlWriter.IndentChar = "`t" 
    $XmlDocument.WriteContentTo($XmlWriter) 
    $XmlWriter.Flush() 
    $StringWriter.Flush() 
    Return $StringWriter.ToString() 
}

<# Function Invoke-RestMethod requires Power Shell v3 or higher #>
if ($PSVersionTable.PSVersion.Major -ge 3) {

    <# Authenticate against Azure AD and retrieve OAuth token #>
    $OauthBody = @{
        grant_type = "client_credentials";
        resource = $Resource;
        client_id = $ClientID;
        client_secret = $ClientSecret
    }
    $OauthUri = "{0}/{1}/oauth2/token?api-version=1.0" -f $LoginURL, $TenantIdentifier
    try {
        $Oauth = Invoke-RestMethod -Method Post -Uri $OauthUri -Body $OauthBody
    } catch {
        New-PrtgError -ErrorText "Error authenticating"
    }

    <# Get service status via REST API #>
    $Operation = "CurrentStatus"
    $ServiceUri = "https://manage.office.com/api/v1.0/{0}/ServiceComms/{1}" -f $TenantIdentifier, $Operation
    $headerParams  = @{
        Authorization = "{0} {1}" -f $oauth.token_type, $oauth.access_token
    }
    try {
        $Data = Invoke-WebRequest -Headers $headerParams -Uri $ServiceUri -UseBasicParsing | ConvertFrom-Json
    } catch {
        New-PrtgError -ErrorText "Error retrieving data"
    }
} else {
    New-PrtgError -ErrorText "PowerShell Version 3.0 or higher required on probe"
}

$Result = @()
foreach ($Item in $Data.value) {
    $Result += @{
        Channel = $Item.WorkloadDisplayName;
        Value = [int]$ReturnCodes.($Item.Status);
        ValueLookup = "custom.office365.state"
    }
}

Out-Prtg -MonitoringData $Result
sleep 0 # VScode debug only