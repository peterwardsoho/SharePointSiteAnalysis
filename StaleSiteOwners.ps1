# Begin try-catch block to handle exceptions
try {
    # Maximum function count
    $maximumfunctioncount = 32768
    
     #Check if Microsoft.Graph module is available, if not install it
   if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
        Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
        Import-Module -Name Microsoft.Graph
         
    }

     #Check if PnP.PowerShell module is available, if not install it
    if (-not (Get-Module -Name PnP.PowerShell -ListAvailable)) {
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        Import-Module -Name PnP.PowerShell
    }
    # Prompt user to confirm global admin status and get tenant name from user
    $Provision = $(Write-Host "Before running this script make sure you are global admin. Type Y if you are and N if you are'nt: " -ForegroundColor Green -NoNewline; Read-Host)

    if($Provision.ToUpper() -eq "Y"){
          $TenantName = $(Write-Host "Enter the Azure Tenant Name e.g contoso: " -ForegroundColor Yellow -NoNewline; Read-Host)

         # Construct admin site URL
         $AdminSiteURL="https://" + $TenantName + "-admin.sharepoint.com/"

         # Prompt user to select script option
           $ScriptOption = $(Write-Host "What Script would you like to run? (A) Stale Sites or (B) Stale Owners or (C) Fetch Classic Sites: " -ForegroundColor Yellow -NoNewline; Read-Host)
    # If user chooses to run Stale Sites script
If($ScriptOption.ToUpper() -eq "A"){
try {
    # Connect to SharePoint Online admin site
    Write-Host "Validating User using PnP login..."
    Connect-PnPOnline -Url $AdminSiteURL -Interactive
}
catch {
    Write-Host "`nError Message: " $_.Exception.Message
    Write-Host "`nError in Line: " $_.InvocationInfo.Line
    Write-Host "`nError in Line Number: "$_.InvocationInfo.ScriptLineNumber
    Exit
}

# Prompt user to enter time period for site analysis
$TimePeriod = $(Write-Host "Please enter your time period. Q(uarterly)/S(emi-Annual)/A(nnual)/C(ustom): " -ForegroundColor Yellow -NoNewline; Read-Host)

# Get site collections
try {
    Write-Host "Fetching sites..."
    $SiteCollections = Get-PnPTenantSite
}
catch {
    Disconnect-PnPOnline
    Write-Host "`nError Message: " $_.Exception.Message
    Write-Host "`nError in Line: " $_.InvocationInfo.Line
    Write-Host "`nError in Line Number: "$_.InvocationInfo.ScriptLineNumber
    Exit
}

# Initialize array to store site owners
$SiteOwners = @()

# Set target date based on user's time period selection
$targetDate = Get-Date
Switch ($TimePeriod.ToUpper()) {
    "Q" { $targetDate = $targetDate.AddMonths(-3) }
    "S" { $targetDate = $targetDate.AddMonths(-6) }
    "A" { $targetDate = $targetDate.AddYears(-1) } 
    "C" {
        $customDateInput = $(Write-Host "Please enter the date in MM/DD/YYYY format: " -ForegroundColor Yellow -NoNewline; Read-Host)
        $targetDate = [datetime]::ParseExact($customDateInput, 'MM/dd/yyyy', $null)
    }
}

# Define function to filter and export sites
function FilterAndExportSites {
    param(
        [string]$Filter,
        [string]$FileName
    )

    $SiteOwners = @()
    # Filter site collections
    $FilteredSites = $SiteCollections | Where-Object { ($_.Title) -match $Filter }
    foreach ($Site in $FilteredSites) {
        if ($Site.LastContentModifiedDate -le $targetDate) {
            $SiteOwners += [PSCustomObject]@{
                'Site Title' = $Site.Title
                'URL' = $Site.Url
                'Last Modified' = $Site.LastContentModifiedDate.ToString("yyyy:MM:dd")
            }
        }
    }

    # Prompt user to export data to CSV
    $ExportCSV = $(Write-Host "Do you want to export this data into CSV file? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host)
    if ($ExportCSV.ToUpper() -eq "Y") {
        # Prompt user to enter CSV file path
        $CSVPath = $(Write-Host "Please enter the file path(Make sure the path is valid e.g C:\Export or C:) " -ForegroundColor Yellow -NoNewline; Read-Host)
        $CSVPathWithFile = $CSVPath + "\" + $FileName
        Write-Output "Exported file to $($CSVPathWithFile)" 
        $SiteOwners | Export-Csv -Path $CSVPathWithFile -NoTypeInformation
        $SiteOwners | Format-Table
    }
    else {
        $SiteOwners | Format-Table
    }
}

# Loop through site collections and identify stale sites
$AlphabetSequence = $(Write-Host "Fetch sites in batches? (A)A-J or (B)K-S or (C)T-Z or D(All): " -ForegroundColor Yellow -NoNewline; Read-Host)
    Switch ($AlphabetSequence.ToUpper()) {
        "A" {
            FilterAndExportSites -Filter '^[A-J]' -FileName "SiteDetails-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            $ContinueKS = $(Write-Host "Would you like to go on for K-S? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host) -eq "Y"
            if($ContinueKS){
                FilterAndExportSites -Filter '^[K-S]' -FileName "SiteDetailsKS-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            }
            $ContinueTZ = $(Write-Host "Would you like to go on for T-Z? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host) -eq "Y"
            if($ContinueTZ){
                FilterAndExportSites -Filter '^[T-Z]' -FileName "SiteDetailsTZ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            }
        }
        "B" {
            FilterAndExportSites -Filter '^[K-S]' -FileName "SiteDetailsKS-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            $ContinueTZ = $(Write-Host "Would you like to go on for T-Z? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host) -eq "Y"
            if($ContinueTZ){
                FilterAndExportSites -Filter '^[T-Z]' -FileName "SiteDetailsTZ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            }
        }
        "C" {
            FilterAndExportSites -Filter '^[T-Z]' -FileName "SiteDetailsTZ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
        }
        "D" {
            FilterAndExportSites -Filter '.' -FileName "SiteDetails-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
        }
    }
Disconnect-PnPOnline
}

# If the user selects the option to run the Stale Owners script
If ($ScriptOption.ToUpper() -eq "B") {
    # Define function to filter and export sites
function FilterAndExportSites {
    param(
        [string]$Filter,
        [string]$FileName
    )
    # Attempt to authenticate and connect to the SharePoint Online Admin Center
    try {
        Write-Host "Validating User using PnP login..."
        Connect-PnPOnline -Url $AdminSiteURL -Interactive
    }
    catch {
        # Catch any errors during the connection attempt and exit the script
        Write-Host "`nError Message: " $_.Exception.Message
        Write-Host "`nError in Line: " $_.InvocationInfo.Line
        Write-Host "`nError in Line Number: "$_.InvocationInfo.ScriptLineNumber
        Exit
    }

    # Fetch all site collections from SharePoint Online
    try {
        Write-Host "Fetching selected sites..."
            $SiteCollections = Get-PnPTenantSite
    }
    catch {
        # If fetching fails, disconnect and show error details, then exit
        Disconnect-PnPOnline
        Write-Host "`nError Message: " $_.Exception.Message
        Write-Host "`nError in Line: " $_.InvocationInfo.Line
        Write-Host "`nError in Line Number: "$_.InvocationInfo.ScriptLineNumber
        Exit
    }
    $SiteOwners = @()
    # Filter site collections
    $FilteredSites = $SiteCollections | Where-Object { ($_.Title) -match $Filter }
    foreach ($Site in $FilteredSites) {
         If($Site.Template -like 'GROUP*') {
              try{
				$Owners = (Get-PnPMicrosoft365GroupOwners -Identity ($Site.GroupId)  | Select -ExpandProperty Email) -join ";"
				}
				Catch
				{
				 $Owners = "$($Site.GroupId) Group Not Found"
                 Write-Host "Group not found for $($Site.Url)"
				}
            }
            Else {
                $Owners = $Site.Owner
            }
            $SiteOwners += New-Object PSObject -Property @{
                'URL' = $Site.Url
                'Owner(s)' = $Owners
            }
    }
    
    # Disconnect from SharePoint Online as it's no longer needed
    Disconnect-PnPOnline

    # Prepare to connect to Microsoft Graph for accessing sign-in data
    try {
        Write-Host "Validating Graph Permissions..."
        Connect-MgGraph -Scopes "AuditLog.Read.All", "User.Read.All"
    }
    catch {
        # Handle any errors during connection attempt to Microsoft Graph
        Disconnect-MgGraph
        Write-Host "`nError Message: " $_.Exception.Message
        Write-Host "`nError in Line: " $_.InvocationInfo.Line
        Write-Host "`nError in Line Number: "$_.InvocationInfo.ScriptLineNumber
        Exit
    }

    # Ask the user for the time period to analyze last sign-in activity
    $TimePeriod = $(Write-Host "Please enter your time period. Q(uarterly)/S(emi-Annual)/A(nnual)/C(ustom): " -ForegroundColor Yellow -NoNewline; Read-Host)

    # Calculate the target date for last sign-in based on user input
    $targetDate = Get-Date
    Switch ($TimePeriod.ToUpper()) {
        "Q" { $targetDate = $targetDate.AddMonths(-3) }
        "S" { $targetDate = $targetDate.AddMonths(-6) }
        "A" { $targetDate = $targetDate.AddYears(-1) }
        "C" {
            # Allow user to specify a custom date
            $customDateInput = $(Write-Host "Please enter the date in MM/DD/YYYY format: " -ForegroundColor Yellow -NoNewline; Read-Host)
            $targetDate = [datetime]::ParseExact($customDateInput, 'MM/dd/yyyy', $null)
        }
    }

    # Initialize an array to hold the last sign-in details of site owners
    $SiteOwnersLastSeen = @()
    $Properties = @(
    'Id','DisplayName','UserPrincipalName','UserType', 'AccountEnabled', 'SignInActivity'   
)
    # For each site owner, determine if they are considered stale based on last sign-in
    ForEach ($Site in $SiteOwners) {
        Write-Host "Checking owner's last login for site: $($Site.'URL')"
        if ($null -eq $Site.'Owner(s)' -or $Site.'Owner(s)' -eq '') {
            $Site.'Owner(s)' = "Owner not found"
       }
            If ($Site.'Owner(s)' -like "*;*") {
                # If there are multiple owners, split the string and check each
                $Owners = $Site.'Owner(s)'.Split(";")
                ForEach ($Owner in $Owners) {
                    $User = Get-MgUser -Filter "mail eq '$Owner'" -Select $Properties
                    if ($User.SignInActivity.LastSignInDateTime -ne $null) {
                        $LastSignInDateTime = [datetime]$User.SignInActivity.LastSignInDateTime
                        if ($LastSignInDateTime -le $targetDate) {
                            # If last sign-in is before the target date, add to report
                            $SiteOwnersLastSeen += New-Object PSObject -Property @{
                                'Owner'      = $Owner
                                'LastSignIn' = $LastSignInDateTime.ToString('yyyy-MM-dd')
                                'URL'        = $Site.'URL'
                            }
                        }
                    }
                }
            }
            elseif ($Site.'Owner(s)' -like "*Not Found*") {
                # Handle cases where the group owner is not found
                $SiteOwnersLastSeen += New-Object PSObject -Property @{
                    'Owner'      = $Site.'Owner(s)'
                    'LastSignIn' = "Not Found"
                    'URL'        = $Site.'URL'
                }
            }
            Else {
              
                # Single owner sites, fetch and check last sign-in date
                $User = Get-MgUser -Filter "mail eq '$($Site.'Owner(s)')'" -Select $Properties
                if ($User.SignInActivity.LastSignInDateTime -ne $null) {
                    $LastSignInDateTime = [datetime]$User.SignInActivity.LastSignInDateTime
                    if ($LastSignInDateTime -le $targetDate) {
                        $SiteOwnersLastSeen += New-Object PSObject -Property @{
                            'Owner'      = $Site.'Owner(s)'
                            'LastSignIn' = $LastSignInDateTime.ToString('yyyy-MM-dd')
                            'URL'        = $Site.'URL'
                        }
                    }
                }
            }
    }

    # Prompt user to export data to CSV
    $ExportCSV = $(Write-Host "Do you want to export this data into CSV file? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host)
    if ($ExportCSV.ToUpper() -eq "Y") {
        # Prompt user to enter CSV file path
        $CSVPath = $(Write-Host "Please enter the file path(Make sure the path is valid e.g C:\Export or C:) " -ForegroundColor Yellow -NoNewline; Read-Host)
        $CSVPathWithFile = $CSVPath + "\" + $FileName
        Write-Output "Exported file to $($CSVPathWithFile)" 
        $SiteOwnersLastSeen | Export-Csv -Path $CSVPathWithFile -NoTypeInformation
        $SiteOwnersLastSeen | Format-Table
    }
    else {
        $SiteOwnersLastSeen | Format-Table
    }
         # Clean up by disconnecting from Microsoft Graph
         Disconnect-MgGraph
}

    $AlphabetSequence = $(Write-Host "Fetch sites in batches? (A)A-J or (B)K-S or (C)T-Z or D(All): " -ForegroundColor Yellow -NoNewline; Read-Host)
    Switch ($AlphabetSequence.ToUpper()) {
        "A" {
            FilterAndExportSites -Filter '^[A-J]' -FileName "SiteOwnersAJ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            $ContinueKS = $(Write-Host "Would you like to go on for K-S? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host) -eq "Y"
            if($ContinueKS){
                FilterAndExportSites -Filter '^[K-S]' -FileName "SiteOwnersKS-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            }
            $ContinueTZ = $(Write-Host "Would you like to go on for T-Z? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host) -eq "Y"
            if($ContinueTZ){
                FilterAndExportSites -Filter '^[T-Z]' -FileName "SiteOwnersTZ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            }
        }
        "B" {
            FilterAndExportSites -Filter '^[K-S]' -FileName "SiteOwnersKS-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            $ContinueTZ = $(Write-Host "Would you like to go on for T-Z? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host) -eq "Y"
            if($ContinueTZ){
                FilterAndExportSites -Filter '^[T-Z]' -FileName "SiteOwnersTZ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            }
        }
        "C" {
            FilterAndExportSites -Filter '^[T-Z]' -FileName "SiteOwnersTZ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
        }
        "D" {
            FilterAndExportSites -Filter '.' -FileName "SiteOwners-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
        }
    }

}

# This PowerShell script is designed for managing SharePoint Online site collections. 
# It allows for the connection to the SharePoint Online Admin site, categorizes site collections 
# based on their title initials (A-J, K-S, T-Z), identifies classic sites using the "STS#0" template, 
# and optionally exports the information to a CSV file.

# Check if the script option selected by the user is to manage site collections ('C' for Collections)
  If($ScriptOption.ToUpper() -eq "C"){
        
        try {
            # Connect to SharePoint Online admin site
            Write-Host "Validating User using PnP login..."
            Connect-PnPOnline -Url $AdminSiteURL -Interactive
            
        }
        catch {
            Write-Host "`nError Message: " $_.Exception.Message
            Write-Host "`nError in Line: " $_.InvocationInfo.Line
            Write-Host "`nError in Line Number: "$_.InvocationInfo.ScriptLineNumber
            Exit
        }

        # Get site collections
        try {
            Write-Host "Fetching sites..."
            $SiteCollections = Get-PnPTenantSite
        }
        catch {
            Disconnect-PnPOnline
            Write-Host "`nError Message: " $_.Exception.Message
            Write-Host "`nError in Line: " $_.InvocationInfo.Line
            Write-Host "`nError in Line Number: "$_.InvocationInfo.ScriptLineNumber
            Exit
        }
        # Define function to filter and export sites
function FilterAndExportSites {
    param(
        [string]$Filter,
        [string]$FileName
    )

    $ClassicSites = @()
    # Filter site collections
    $FilteredSites = $SiteCollections | Where-Object { ($_.Title) -match $Filter }
    foreach ($Site in $FilteredSites) {
        if($Site.Template -eq "STS#0"){
                $ClassicSites += New-Object PSObject -Property @{
                    'Site Title' = $Site.Title
                    'URL' = $Site.Url
                    'Last Modified' = $Site.LastContentModifiedDate.ToString("yyyy:MM:dd")
                    'Site Template' = $Site.Template
                }
            }
    }

    # Prompt user to export data to CSV
    $ExportCSV = $(Write-Host "Do you want to export this data into CSV file? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host)
    if ($ExportCSV.ToUpper() -eq "Y") {
        # Prompt user to enter CSV file path
        $CSVPath = $(Write-Host "Please enter the file path(Make sure the path is valid e.g C:\Export or C:) " -ForegroundColor Yellow -NoNewline; Read-Host)
        $CSVPathWithFile = $CSVPath + "\" + $FileName
        Write-Output " file to $($CSVPathWithFile)" 
        $ClassicSites | Export-Csv -Path $CSVPathWithFile -NoTypeInformation
         $ClassicSites | Format-Table
    }
    else {
        $ClassicSites | Format-Table
    }
}
        $AlphabetSequence = $(Write-Host "Fetch sites in batches? (A)A-J or (B)K-S or (C)T-Z or D(All): " -ForegroundColor Yellow -NoNewline; Read-Host)
        # Loop through site collections and identify stale sites
Switch ($AlphabetSequence.ToUpper()) {
        "A" {
            FilterAndExportSites -Filter '^[A-J]' -FileName "ClassicSitesAJ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            $ContinueKS = $(Write-Host "Would you like to go on for K-S? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host) -eq "Y"
            if($ContinueKS){
                FilterAndExportSites -Filter '^[K-S]' -FileName "ClassicSitesKS-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            }
            $ContinueTZ = $(Write-Host "Would you like to go on for T-Z? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host) -eq "Y"
            if($ContinueTZ){
                FilterAndExportSites -Filter '^[T-Z]' -FileName "ClassicSitesTZ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            }
        }
        "B" {
            FilterAndExportSites -Filter '^[K-S]' -FileName "ClassicSitesKS-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            $ContinueTZ = $(Write-Host "Would you like to go on for T-Z? Y(es) or N(o): " -ForegroundColor Yellow -NoNewline; Read-Host) -eq "Y"
            if($ContinueTZ){
                FilterAndExportSites -Filter '^[T-Z]' -FileName "ClassicSitesTZ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
            }
        }
        "C" {
            FilterAndExportSites -Filter '^[T-Z]' -FileName "ClassicSitesTZ-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
        }
        "D" {
            FilterAndExportSites -Filter '.' -FileName "ClassicSites-$($TenantName)-$((Get-Date).ToString('MM-dd-yyyy')).csv"
        }
    }
        Disconnect-PnPOnline
    }

   } 
}

catch {
    # Catch any errors that occur during script execution
    Write-Host "`nError Message: " $_.Exception.Message
    Write-Host "`nError in Line: " $_.InvocationInfo.Line
    Write-Host "`nError in Line Number: "$_.InvocationInfo.ScriptLineNumber
    Exit
}
finally {
    # Clear any errors that may have occurred
    $error.clear()
}

