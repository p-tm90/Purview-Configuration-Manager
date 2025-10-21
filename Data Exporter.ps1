#Settings script execution location.
$Location = $PSScriptRoot
Set-Location $Location
$OutputsPath = "$Location\Outputs"

#Importing core functions
. .\Dependencies\Functions.ps1

#Install & update relevant Purview modules.
Configure-MicrosoftModules -TypeOfConfiguration Purview -InstallScope CurrentUser

#Core script pre-requisites setup.
$Time = (Get-Date).ToString('yyyy-MM-dd-HHmmss')

#Information gathering from Purview
Write-Host "Gathering sensitivity label configuration from Purview."
$Labels = Get-Label
Write-Host "Gathering sensitive label publishing policy configuration from Purview."
$LabelPolicies = Get-LabelPolicy
Write-Host "Gathering automatic labeling policy configuration from Purview."
$AutoLabelPolicies = Get-AutoSensitivityLabelPolicy
Write-Host "Gathering automatic labeling policy rules configuration from Purview."
$AutoLabelRules = Get-AutoSensitivityLabelRule
Write-Host "Gathering DLP policies configuration from Purview."
$DLPPolicies = Get-DLPCompliancePolicy
Write-Host "Gathering DLP rules configuration from Purview."
$DLPRules = Get-DLPComplianceRule
Write-Host "Gathering DLP info types configuration from Purview."
$DLPSITs = Get-DLPSensitiveInformationType

#Processing raw information from Purview into readable format with highlights.
$LabelInfo = @()
Foreach ($Label in $Labels){
    $LabelInfo += [PSCustomObject]@{
        GUID = $Label.GUID
        DisplayName = $Label.DisplayName
        Name = $Label.Name
        Tooltip = $Label.Tooltip
        ParentID = $Label.ParentID
        IsParent = $Label.IsParent
        Description = $Label.Description
        Mode = $Label.Mode
        ContentType = $Label.ContentType
        CreatedBy = $Label.CreatedBy
        LastModifiedBy = $Label.LastModifiedBy
    }
}

$LabelPolicyInfo = @()
Foreach ($LabelPolicy in $LabelPolicies){
    $LabelPolicyInfo += [PSCustomObject]@{
        GUID = $LabelPolicy.GUID
        Name = $LabelPolicy.Name
        Labels = $LabelPolicy.Labels
        Comment = $LabelPolicy.Comment
        Enabled = $LabelPolicy.Enabled
        Mode = $LabelPolicy.Mode
        CreatedBy = $LabelPolicy.CreatedBy
        LastModifiedBy = $LabelPolicy.LastModifiedBy
    }
}

$AutoLabelPolicyInfo = @()
Foreach ($AutoLabelPolicy in $AutoLabelPolicies){
    $AutoLabelPolicyInfo += [PSCustomObject]@{
        LabelDisplayName = $AutoLabelPolicy.LabelDisplayName
        Name = $AutoLabelPolicy.Name
        Type = $AutoLabelPolicy.Type
        OverwriteLabel = $AutoLabelPolicy.OverwriteLabel
        Comment = $AutoLabelPolicy.Comment
        Enabled = $AutoLabelPolicy.Enabled
        Mode = $AutoLabelPolicy.Mode
        Workload = $AutoLabelPolicy.Workload
        ContentType = $AutoLabelPolicy.ContentType
        CreatedBy = $AutoLabelPolicy.CreatedBy
        LastModifiedBy = $AutoLabelPolicy.LastModifiedBy
    }
}

$AutoLabelRuleInfo = @()
foreach ($AutoLabelRule in $AutoLabelRules){
    $AutoLabelRuleInfo += [PSCustomObject]@{
        ParentPolicyName = $AutoLabelRule.ParentPolicyName
        DisplayName = $AutoLabelRule.DisplayName
        ReportSeverityLevel = $AutoLabelRule.ReportSeverityLevel
        Comment = $AutoLabelRule.Comment
        Disabled = $AutoLabelRule.Disabled
        Mode = $AutoLabelRule.Mode
        Workload = $AutoLabelRule.Workload
        ContentType = $AutoLabelRule.ContentType
        CreatedBy = $AutoLabelRule.CreatedBy
        LastModifiedBy = $AutoLabelRule.LastModifiedBy
    }
}

$DLPPolicyInfo = @()
Foreach ($DLPPolicy in $DLPPolicies){
    $DLPPolicyInfo += [PSCustomObject]@{
        Name = $DLPPolicy.Name
        DisplayName = $DLPPolicy.DisplayName
        IsSimulationMode = $DLPPolicy.IsSimulationMode
        Workload = $DLPPolicy.Workload
        Enabled = $DLPPolicy.Enabled
        Comment = $DLPPolicy.Comment
        CreatedBy = $DLPPolicy.CreatedBy
        LastModifiedBy = $DLPPolicy.LastModifiedBy
    }
}

$DLPRuleInfo = @()
Foreach ($DLPRule in $DLPRules){
    $DLPRuleInfo += [PSCustomObject]@{
        GUID = $DLPRule.GUID
        DisplayName = $DLPRule.DisplayName
        ParentPolicyName = $DLPRule.ParentPolicyName
        ReportSeverityLevel = $DLPRule.ReportSeverityLevel
        Workload = $DLPRule.Workload
        Disabled = $DLPRule.Disabled
        Mode = $DLPRule.Mode
        Comment = $DLPRule.Comment
        CreatedBy = $DLPRule.CreatedBy
        LastModifiedBy = $DLPRule.LastModifiedBy
    }
}

$DLPSITInfo = @()
Foreach ($DLPSIT in $DLPSITs){
    $DLPSITInfo += [PSCustomObject]@{
        Name = $DLPSIT.Name
        Description = $DLPSIT.Description
        Publisher = $DLPSIT.Publisher
        Type = $DLPSIT.Type
    }
}

#Extract compiled information into Excel and TXT output files.
Write-Host "Exporting raw data to outputs folder path as separate TXT files per configured component."
$FilePath = "$($OutputsPath)\$($Time)_RawData"
$Labels | ForEach-Object {$_ | Select-Object * | Out-File -FilePath "$($FilePath)_Labels_$($_.GUID).txt"}
$LabelPolicies | ForEach-Object {$_ | Select-Object * | Out-File -FilePath "$($FilePath)_LabelPolicies_$($_.GUID).txt"}

Foreach ($a in $AutoLabelPolicies){
    New-Item -ItemType Directory -Path $OutputsPath\AutoLabelPolicies\$($a.GUID)
    $NestedFilePath = "$($OutputsPath)\AutoLabelPolicies\$($a.GUID)\$($Time)_RawData"
    $a | Select-Object * | Out-File -FilePath "$($NestedFilePath)_AutoLabelPolicies_$($a.GUID).txt"
    Foreach ($b in $AutoLabelRules){
        If ($b.ParentPolicyName -eq "$($a.Name)"){$b | Select-Object * | Out-File "$($NestedFilePath)_AutoLabelRules_$($b.GUID).txt"}
    }
}

Foreach ($a in $DLPPolicies){
    New-Item -ItemType Directory -Path $OutputsPath\DLPPolicies\$($a.GUID)
    $NestedFilePath = "$($OutputsPath)\DLPPolicies\$($a.GUID)\$($Time)_RawData"
    $a | Select-Object * | Out-File -FilePath "$($NestedFilePath)_DLPPolicies_$($a.GUID).txt"
    Foreach ($b in $DLPRules){
        If ($b.ParentPolicyName -eq "$($a.Name)"){$b | Select-Object * | Out-File "$($NestedFilePath)_DLPRules_$($b.GUID).txt"}
    }
}

$DLPSITs | ForEach-Object {$_ | Select-Object * | Out-File -FilePath "$($FilePath)_DLPSITs_$($_.GUID).txt"}

#Export processed information to Excel outputs file.
Write-Host "Exporting processed data to outputs folder path as single XLSX file."
$FilePath = "$($OutputsPath)\$($Time)_ProcessingInfoData.xlsx"
$LabelInfo = Export-Excel -Path $FilePath -WorksheetName "Sensitivity Labels"
$LabelPolicyInfo = Export-Excel -Path $FilePath -WorksheetName "Label Policies"
$AutoLabelPolicyInfo = Export-Excel -Path $FilePath -WorksheetName "AutoLabel Policies"
$AutoLabelRuleInfo = Export-Excel -Path $FilePath -WorksheetName "AutoLabel Rules"
$DLPPolicyInfo = Export-Excel -Path $FilePath -WorksheetName "DLP Policies"
$DLPRuleInfo = Export-Excel -Path $FilePath -WorksheetName "DLP Rules"
$DLPSITInfo = Export-Excel -Path $FilePath -WorksheetName "DLP SITs"

#Completion message
Write-Host 'Export has completed!'