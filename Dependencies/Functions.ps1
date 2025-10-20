function Connect-MicrosoftModules {
    param (
        [Parameter(Mandatory="True",Position=0)]
        [ValidateSet("Purview","All")]
        [string]$TypeOfRun
    )
    #Header text for installed module check & connection attempt.
    $HeaderText = "Initiating connection attempts to desired modules... NOTE: Authentication window may pop-under IDE in use when running this script!"
    $HeaderTextDivider = "="*$HeaderText.length
    Write-Host "$($HeaderText)`n$($HeaderTextDivider)" -ForegroundColor Cyan

    #Gather currently installed modules and sort by relevant-to-M365 items.
    $InstalledModule = Get-InstalledModule
    $TargetModules = $InstalledModules | Where-Object {
        $_.Name -like "*ExchangeOnlineManagement*" -OR `
        $_.Name -like "*SharePoint.Online*" -OR `
        $_.Name -like "*Microsoft.Graph*" -OR `
        $_.Name -like "*AzureAD*"
    }

    If (!$TypeOfRun){
        Write-Host "No configured connection type specified (either 'Purview' or 'All' values currently supported). Shutting down function, please re-run with specified parameters." -ForegroundColor Red
        exit
    }
    Elseif ($TypeOfRun -eq "Purview"){
        Import-Module ExchangeOnlineManagement
        $PurviewCheck = Get-Command Get-DLPSensitiveInformationTypeConfig -ErrorAction SilentlyContinue
        If (!$PurviewCheck){
            Write-Host "Please enter the UserPrincipalName (UPN) value you would use to sign into an account for administration within Microsoft Purview`
            (e.g., 'ADMIN_John.Smith@domain.com')" -ForegroundColor Yellow
            $AdminUPN = Read-Host "Enter UPN here"
            Connect-IPPSSession}
    }
    Elseif ($TypeOfRun -eq "All"){
        Foreach ($Module in $TargetModules){
            Import-Module -name $Module.Name
        }
        Write-Host "Please enter the built-in subdomain associated with your SharePoint online tenancy.`
        (E.g., Enter 'DemoDomain' if your SharePoint instance displays https://DemoDomain.sharepoint.com/sites/SITENAME when navigating through various sites." -ForegroundColor Yellow
        $URIBase = Read-Host
        $URI = "https://"+"$($URIBase)"+"-admin.sharepoint.com"
    }
    
}