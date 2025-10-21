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
            Connect-IPPSSession -UserPrincipalName $AdminUPN
        }
    }
    Elseif ($TypeOfRun -eq "All"){
        Foreach ($Module in $TargetModules){
            Import-Module -name $Module.Name
        }
        Write-Host "Please enter the built-in subdomain associated with your SharePoint online tenancy.`
        (E.g., Enter 'DemoDomain' if your SharePoint instance displays https://DemoDomain.sharepoint.com/sites/SITENAME when navigating through various sites." -ForegroundColor Yellow
        $URIBase = Read-Host
        $URI = "https://"+"$($URIBase)"+"-admin.sharepoint.com"

        #Connect to Purview
        $PurviewCheck = Get-Command Get-DLPSensitiveInformationTypeConfig -ErrorAction SilentlyContinue
        If (!$PurviewCheck){
            Write-Host "Please enter the UserPrincipalName (UPN) value you would use to sign into an account for administration within Microsoft Purview`
            (e.g., 'ADMIN_John.Smith@domain.com')" -ForegroundColor Yellow
            $AdminUPN = Read-Host "Enter UPN here"
            Connect-IPPSSession
        }
        $AADCheck = Get-Command Get-AzureADUser
        If (!$AADCheck){Connect-AzureAD}

        $MgGraphCheck = Get-Command Get-MgUser
        If (!$MgGraphCheck){Connect-MgGraph -Scopes "User.Read.All","Group.Read.All"}

        $SPOCheck = Get-Command Get-SPOSite
        If (!$SPOCheck){Connect-SPOService -URL $URI}
    }
}

function Configure-MicrosoftModules{
    param (
        [Parameter(Mandatory="True",Position="0")]
        [ValidateSet("Purview","All")]
        [string]$TypeOfConfiguration,
        [Parameter(Mandatory="True",Position="0")]
        [ValidateSet("CurrentUser","Device")]
        [String]$InstallScope
    )

    #Capture current installed module list, then compare against web modules and update as needed.
    $HeaderText = "Gathering list of installed modules in preparation for module updates and/or installs."
    $HeaderTextDivider = "="*$HeaderText.Length
    $InstalledModules = Get-InstalledModule

    If ($TypeOfConfiguration -eq "Purview"){
        $WebModuleNames = @("ExchangeOnlineManagement","ImportExcel")
        $WebModules = @()
        $WebModuleNames | Foreach-Object {
            Write-Host "Checking for $_ module's most current version online." -ForegroundColor Cyan
            $WebModules += Find-Module -Name $_
        }
    }

    If ($TypeOfConfiguration -eq "All"){
        $WebModuleNames = @("ExchangeOnlineManagement","Microsoft.Graph","Microsoft.Online.SharePoint.PowerShell","ImportExcel")
        $WebModules = @()
        $WebModuleNames | Foreach-Object {
            Write-Host "Checking for $_ modules most current version online." -ForegroundColor Cyan
            $WebModules += Find-Module -Name $_
        }
    }

    Foreach ($a in $WebModules){
        Write-Host "Comparing $($a.Name) module's versions on both installed and online instances." -ForegroundColor Cyan
        $CurrentModule = $InstalledModules | ?{$_.Name -eq $a.Name}

        #If current installed module version is mismatched...
        If($CurrentModule.Version -ne $a.Version){
            Write-Host "`nInstalled module $($a.Name) is missing or our of date!`nCurrent installed version is: $($CurrentModule.Version)`nWeb version is: $($a.Version)" -ForegroundColor Red
            If (!$CurrentModule){
                Write-Host "Installed module $($a.Name)..."
                Install-Module -Name $($a.Name) -Scope $InstallScope
            }
            Elseif ($CurrentModule){
                Write-Host "Updating module $($a.Name)..."
                Update-Module -Name $($a.Name) -Scope $InstallScope
            }
        }
        Elseif($CurrentModule.Version -eq $a.Version){
            Write-Host "`nInstalled module $($a.Name) is up-to-date! Moving to next module." -ForegroundColor Green
        }
        Else {
            Write-Host "`nUnable to locate $($a.Name) in current installed modules list, running new install as default action..." -ForegroundColor Red
            Install-Module -Name $($a.Name) -Scope $InstallScope
        }
    }
    Connect-MicrosoftModules -TypeOfRun $TypeOfConfiguration
}