<#
$Credential = Get-Credential
$ExchangeSession = New-PSSession `
    -ConfigurationName Microsoft.Exchange `
    -ConnectionUri "https://outlook.office365.com/powershell-liveid" `
    -Credential $credential -Authentication "Basic" -AllowRedirection
    Import-PSSession $ExchangeSession
#>

$start_date = (get-date).AddDays(-7).ToString('MM/dd/yyy')
$end_date = (get-date).ToString('MM/dd/yyy')
$filestamp = (get-date).ToString('yyyyMMdd')

$results = Search-UnifiedAuditLog -StartDate $start_date -EndDate $end_date `
    -SessionCommand ReturnLargeSet `
    -ResultSize 1000 `
    #-RecordType AzureActiveDirectoryAccountLogon
$results | Out-File result.txt

$oUnifiedAuditLogs = @()
$oAzureActiveDirectorys = @()
$oAzureActiveDirectoryAccountLogons = @()
$oAzureActiveDirectoryStsLogons = @()
$oComplianceDLPExchanges = @()
$oComplianceDLPSharePoints = @()
$oCRMs = @()
$oDataCenterSecurityCmdlets = @()
$oDiscoverys = @()
$oExchangeAdmins = @()
$oExchangeAggregatedOperations = @()
$oExchangeItems = @()
$oExchangeItemGroups = @()
$oMicrosoftTeamss = @()
$oMicrosoftTeamsAddOnss = @()
$oMicrosoftTeamsSettingsOperations = @()
$oOneDrives = @()
$oPowerBIAudits = @()
$oSecurityComplianceCenterEOPCmdlets = @()
$oSharePoints = @()
$oSharePointFileOperations = @()
$oSharePointSharingOperations = @()
$oSkypeForBusinessCmdletss = @()
$oSkypeForBusinessPSTNUsages = @()
$oSkypeForBusinessUsersBlockeds = @()
$oSways = @()
$oThreatIntelligences = @()
$oYammers = @()

foreach ($result in $results) {
    $oUnifiedAuditLog = New-Object –TypeName PSObject
    $oUnifiedAuditLog | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
    $oUnifiedAuditLog | Add-Member –MemberType NoteProperty –Name RunspaceId –Value $result.RunspaceId
    $oUnifiedAuditLog | Add-Member –MemberType NoteProperty –Name RecordType –Value $result.RecordType
    $oUnifiedAuditLog | Add-Member –MemberType NoteProperty –Name CreationDate –Value $result.CreationDate
    $oUnifiedAuditLog | Add-Member –MemberType NoteProperty –Name UserIds –Value $result.UserIds
    $oUnifiedAuditLog | Add-Member –MemberType NoteProperty –Name ResultIndex –Value $result.ResultIndex
    $oUnifiedAuditLog | Add-Member –MemberType NoteProperty –Name ResultCount –Value $result.ResultCount
    $oUnifiedAuditLog | Add-Member –MemberType NoteProperty –Name IsValid –Value $result.IsValid
    $oUnifiedAuditLog | Add-Member –MemberType NoteProperty –Name ObjectState –Value $results.ObjectState
    $oUnifiedAuditLogs += $oUnifiedAuditLog
    
    if ($result.RecordType -eq 'PowerBIAudit') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oItem += $oItem
    }

    elseif ($result.RecordType -eq 'AzureActiveDirectoryAccountLogon') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oItem | Add-Member –MemberType NoteProperty –Name ExtendedPropertiesName –Value $result.AuditData.ExtendedProperties[0].Name
        $oItem | Add-Member –MemberType NoteProperty –Name ExtendedPropertiesValue –Value $result.AuditData.ExtendedProperties[0].Value 
        $oAzureActiveDirectoryAccountLogons += $oItem
    }

    elseif ($result.RecordType -eq 'ExchangeItem') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oExchangeItems += $oItem
    }
}


$oAzureActiveDirectory | Export-Csv -NoTypeInformation "$filestamp-AzureActiveDirectory.csv"
$oAzureActiveDirectoryAccountLogon | Export-Csv -NoTypeInformation "$filestamp-AzureActiveDirectoryAccountLogon.csv"
$oAzureActiveDirectoryStsLogon | Export-Csv -NoTypeInformation "$filestamp-AzureActiveDirectoryStsLogon.csv"
$oComplianceDLPExchange | Export-Csv -NoTypeInformation "$filestamp-ComplianceDLPExchange.csv"
$oComplianceDLPSharePoint | Export-Csv -NoTypeInformation "$filestamp-ComplianceDLPSharePoint.csv"
$oCRM | Export-Csv -NoTypeInformation "$filestamp-CRM.csv"
$oDataCenterSecurityCmdlet | Export-Csv -NoTypeInformation "$filestamp-DataCenterSecurityCmdlet.csv"
$oDiscovery | Export-Csv -NoTypeInformation "$filestamp-Discovery.csv"
$oExchangeAdmin | Export-Csv -NoTypeInformation "$filestamp-ExchangeAdmin.csv"
$oExchangeAggregatedOperation | Export-Csv -NoTypeInformation "$filestamp-ExchangeAggregatedOperation.csv"
$oExchangeItem | Export-Csv -NoTypeInformation "$filestamp-ExchangeItem.csv"
$oExchangeItemGroup | Export-Csv -NoTypeInformation "$filestamp-ExchangeItemGroup.csv"
$oMicrosoftTeams | Export-Csv -NoTypeInformation "$filestamp-MicrosoftTeams.csv"
$oMicrosoftTeamsAddOns | Export-Csv -NoTypeInformation "$filestamp-MicrosoftTeamsAddOns.csv"
$oMicrosoftTeamsSettingsOperation | Export-Csv -NoTypeInformation "$filestamp-MicrosoftTeamsSettingsOperation.csv"
$oOneDrive | Export-Csv -NoTypeInformation "$filestamp-OneDrive.csv"
$oPowerBIAudit | Export-Csv -NoTypeInformation "$filestamp-PowerBIAudit.csv"
$oSecurityComplianceCenterEOPCmdlet | Export-Csv -NoTypeInformation "$filestamp-SecurityComplianceCenterEOPCmdlet.csv"
$oSharePoint | Export-Csv -NoTypeInformation "$filestamp-SharePoint.csv"
$oSharePointFileOperation | Export-Csv -NoTypeInformation "$filestamp-SharePointFileOperation.csv"
$oSharePointSharingOperation | Export-Csv -NoTypeInformation "$filestamp-SharePointSharingOperation.csv"
$oSkypeForBusinessCmdlets | Export-Csv -NoTypeInformation "$filestamp-SkypeForBusinessCmdlets.csv"
$oSkypeForBusinessPSTNUsage | Export-Csv -NoTypeInformation "$filestamp-SkypeForBusinessPSTNUsage.csv"
$oSkypeForBusinessUsersBlocked | Export-Csv -NoTypeInformation "$filestamp-SkypeForBusinessUsersBlocked.csv"
$oSway | Export-Csv -NoTypeInformation "$filestamp-Sway.csv"
$oThreatIntelligence | Export-Csv -NoTypeInformation "$filestamp-ThreatIntelligence.csv"
$oYammer | Export-Csv -NoTypeInformation "$filestamp-Yammer.csv"

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
