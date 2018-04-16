<#
$Credential = Get-Credential
$ExchangeSession = New-PSSession `
    -ConfigurationName Microsoft.Exchange `
    -ConnectionUri "https://outlook.office365.com/powershell-liveid" `
    -Credential $credential -Authentication "Basic" -AllowRedirection
    Import-PSSession $ExchangeSession
#>

$filestamp = (get-date).ToString('yyyyMMddhhmmss')

$results = Search-UnifiedAuditLog -StartDate (get-date).AddMinutes(-60) -EndDate (get-date) -SessionCommand ReturnLargeSet -ResultSize 5000
    #-RecordType AzureActiveDirectoryAccountLogon
$results | Out-File result.txt

$UnifiedAuditLog = @()
$AzureActiveDirectory = @()
$AzureActiveDirectoryAccountLogon = @()
$AzureActiveDirectoryStsLogon = @()
$ComplianceDLPExchange = @()
$ComplianceDLPSharePoint = @()
$CRM = @()
$DataCenterSecurityCmdlet = @()
$Discovery = @()
$ExchangeAdmin = @()
$ExchangeAggregatedOperation = @()
$ExchangeItem = @()
$ExchangeItemGroup = @()
$MicrosoftTeams = @()
$MicrosoftTeamsAddOns = @()
$MicrosoftTeamsSettingsOperation = @()
$OneDrive = @()
$PowerBIAudit = @()
$SecurityComplianceCenterEOPCmdlet = @()
$SharePoint = @()
$SharePointFileOperation = @()
$SharePointSharingOperation = @()
$SkypeForBusinessCmdlets = @()
$SkypeForBusinessPSTNUsage = @()
$SkypeForBusinessUsersBlocked = @()
$Sway = @()
$ThreatIntelligence = @()
$Yammer = @()

foreach ($result in $results) {
    $oItem = New-Object –TypeName PSObject
    $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
    $oItem | Add-Member –MemberType NoteProperty –Name RunspaceId –Value $result.RunspaceId
    $oItem | Add-Member –MemberType NoteProperty –Name RecordType –Value $result.RecordType
    $oItem | Add-Member –MemberType NoteProperty –Name CreationDate –Value $result.CreationDate
    $oItem | Add-Member –MemberType NoteProperty –Name UserIds –Value $result.UserIds
    $oItem | Add-Member –MemberType NoteProperty –Name ResultIndex –Value $result.ResultIndex
    $oItem | Add-Member –MemberType NoteProperty –Name ResultCount –Value $result.ResultCount
    $oItem | Add-Member –MemberType NoteProperty –Name IsValid –Value $result.IsValid
    $oItem | Add-Member –MemberType NoteProperty –Name ObjectState –Value $results.ObjectState
    $UnifiedAuditLog += $oItem
    
    if ($result.RecordType -eq 'AzureActiveDirectory') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $AzureActiveDirectory += $oItem 
    }

    elseif ($result.RecordType -eq 'AzureActiveDirectoryAccountLogon') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oItem | Add-Member –MemberType NoteProperty –Name ExtendedPropertiesName –Value $result.AuditData.ExtendedProperties[0].Name
        $oItem | Add-Member –MemberType NoteProperty –Name ExtendedPropertiesValue –Value $result.AuditData.ExtendedProperties[0].Value 
        $AzureActiveDirectoryAccountLogon += $oItem  
    }
 
    elseif ($result.RecordType -eq 'AzureActiveDirectoryStsLogon') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $AzureActiveDirectoryStsLogon += $oItem  
    }
 
    elseif ($result.RecordType -eq 'ComplianceDLPExchange') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $ComplianceDLPExchange += $oItem  
    }
 
    elseif ($result.RecordType -eq 'ComplianceDLPSharePoint') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $ComplianceDLPSharePoint += $oItem  
    }
 
    elseif ($result.RecordType -eq 'CRM') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $CRM += $oItem  
    }
 
    elseif ($result.RecordType -eq 'DataCenterSecurityCmdlet') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $DataCenterSecurityCmdlet += $oItem  
    }
 
    elseif ($result.RecordType -eq 'Discovery') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $Discovery += $oItem  
    }
 
    elseif ($result.RecordType -eq 'ExchangeAdmin') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $ExchangeAdmin += $oItem  
    }
 
    elseif ($result.RecordType -eq 'ExchangeAggregatedOperation') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $ExchangeAggregatedOperation += $oItem  
    }
 
    elseif ($result.RecordType -eq 'ExchangeItem') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $ExchangeItem += $oItem  
    }
 
    elseif ($result.RecordType -eq 'ExchangeItemGroup') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $ExchangeItemGroup += $oItem  
    }
 
    elseif ($result.RecordType -eq 'MicrosoftTeams') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $MicrosoftTeams += $oItem  
    }
 
    elseif ($result.RecordType -eq 'MicrosoftTeamsAddOns') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $MicrosoftTeamsAddOns += $oItem  
    }
 
    elseif ($result.RecordType -eq 'MicrosoftTeamsSettingsOperation') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $MicrosoftTeamsSettingsOperation += $oItem  
    }
 
    elseif ($result.RecordType -eq 'OneDrive') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $OneDrive += $oItem  
    }
 
    elseif ($result.RecordType -eq 'PowerBIAudit') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $PowerBIAudit += $oItem  
    }
 
    elseif ($result.RecordType -eq 'SecurityComplianceCenterEOPCmdlet') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $SecurityComplianceCenterEOPCmdlet += $oItem  
    }
 
    elseif ($result.RecordType -eq 'SharePoint') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $SharePoint += $oItem  
    }
 
    elseif ($result.RecordType -eq 'SharePointFileOperation') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $SharePointFileOperation += $oItem  
    }
 
    elseif ($result.RecordType -eq 'SharePointSharingOperation') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $SharePointSharingOperation += $oItem  
    }
 
    elseif ($result.RecordType -eq 'SkypeForBusinessCmdlets') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $SkypeForBusinessCmdlets += $oItem  
    }
 
    elseif ($result.RecordType -eq 'SkypeForBusinessPSTNUsage') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $SkypeForBusinessPSTNUsage += $oItem  
    }
 
    elseif ($result.RecordType -eq 'SkypeForBusinessUsersBlocked') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $SkypeForBusinessUsersBlocked += $oItem  
    }
 
    elseif ($result.RecordType -eq 'Sway') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $Sway += $oItem  
    }
 
    elseif ($result.RecordType -eq 'ThreatIntelligence') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $ThreatIntelligence += $oItem  
    }
 
    elseif ($result.RecordType -eq 'Yammer') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $Yammer += $oItem  
    }
}

$UnifiedAuditLog | Export-Csv -NoTypeInformation "$filestamp-UnifiedAuditLog.csv"
$AzureActiveDirectory | Export-Csv -NoTypeInformation "$filestamp-AzureActiveDirectory.csv"
$AzureActiveDirectoryAccountLogon | Export-Csv -NoTypeInformation "$filestamp-AzureActiveDirectoryAccountLogon.csv"
$AzureActiveDirectoryStsLogon | Export-Csv -NoTypeInformation "$filestamp-AzureActiveDirectoryStsLogon.csv"
$ComplianceDLPExchange | Export-Csv -NoTypeInformation "$filestamp-ComplianceDLPExchange.csv"
$ComplianceDLPSharePoint | Export-Csv -NoTypeInformation "$filestamp-ComplianceDLPSharePoint.csv"
$CRM | Export-Csv -NoTypeInformation "$filestamp-CRM.csv"
$DataCenterSecurityCmdlet | Export-Csv -NoTypeInformation "$filestamp-DataCenterSecurityCmdlet.csv"
$Discovery | Export-Csv -NoTypeInformation "$filestamp-Discovery.csv"
$ExchangeAdmin | Export-Csv -NoTypeInformation "$filestamp-ExchangeAdmin.csv"
$ExchangeAggregatedOperation | Export-Csv -NoTypeInformation "$filestamp-ExchangeAggregatedOperation.csv"
$ExchangeItem | Export-Csv -NoTypeInformation "$filestamp-ExchangeItem.csv"
$ExchangeItemGroup | Export-Csv -NoTypeInformation "$filestamp-ExchangeItemGroup.csv"
$MicrosoftTeams | Export-Csv -NoTypeInformation "$filestamp-MicrosoftTeams.csv"
$MicrosoftTeamsAddOns | Export-Csv -NoTypeInformation "$filestamp-MicrosoftTeamsAddOns.csv"
$MicrosoftTeamsSettingsOperation | Export-Csv -NoTypeInformation "$filestamp-MicrosoftTeamsSettingsOperation.csv"
$OneDrive | Export-Csv -NoTypeInformation "$filestamp-OneDrive.csv"
$PowerBIAudit | Export-Csv -NoTypeInformation "$filestamp-PowerBIAudit.csv"
$SecurityComplianceCenterEOPCmdlet | Export-Csv -NoTypeInformation "$filestamp-SecurityComplianceCenterEOPCmdlet.csv"
$SharePoint | Export-Csv -NoTypeInformation "$filestamp-SharePoint.csv"
$SharePointFileOperation | Export-Csv -NoTypeInformation "$filestamp-SharePointFileOperation.csv"
$SharePointSharingOperation | Export-Csv -NoTypeInformation "$filestamp-SharePointSharingOperation.csv"
$SkypeForBusinessCmdlets | Export-Csv -NoTypeInformation "$filestamp-SkypeForBusinessCmdlets.csv"
$SkypeForBusinessPSTNUsage | Export-Csv -NoTypeInformation "$filestamp-SkypeForBusinessPSTNUsage.csv"
$SkypeForBusinessUsersBlocked | Export-Csv -NoTypeInformation "$filestamp-SkypeForBusinessUsersBlocked.csv"
$Sway | Export-Csv -NoTypeInformation "$filestamp-Sway.csv"
$ThreatIntelligence | Export-Csv -NoTypeInformation "$filestamp-ThreatIntelligence.csv"
$Yammer | Export-Csv -NoTypeInformation "$filestamp-Yammer.csv"

<#
AzureActiveDirectory
AzureActiveDirectoryAccountLogon
AzureActiveDirectoryStsLogon
ComplianceDLPExchange
ComplianceDLPSharePoint
CRM
DataCenterSecurityCmdlet
Discovery
ExchangeAdmin
ExchangeAggregatedOperation
ExchangeItem
ExchangeItemGroup
MicrosoftTeams
MicrosoftTeamsAddOns
MicrosoftTeamsSettingsOperation
OneDrive
PowerBIAudit
SecurityComplianceCenterEOPCmdlet
SharePoint
SharePointFileOperation
SharePointSharingOperation
SkypeForBusinessCmdlets
SkypeForBusinessPSTNUsage
SkypeForBusinessUsersBlocked
Sway
ThreatIntelligence
Yammer
#>