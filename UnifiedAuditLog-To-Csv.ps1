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


$oAzureActiveDirectorys | Export-Csv -NoTypeInformation "$filestamp-AzureActiveDirectory.csv"
$oAzureActiveDirectoryAccountLogons | Export-Csv -NoTypeInformation "$filestamp-AzureActiveDirectoryAccountLogon.csv"
$oAzureActiveDirectoryStsLogons | Export-Csv -NoTypeInformation "$filestamp-AzureActiveDirectoryStsLogon.csv"
$oComplianceDLPExchanges | Export-Csv -NoTypeInformation "$filestamp-ComplianceDLPExchange.csv"
$oComplianceDLPSharePoints | Export-Csv -NoTypeInformation "$filestamp-ComplianceDLPSharePoint.csv"
$oCRMs | Export-Csv -NoTypeInformation "$filestamp-CRM.csv"
$oDataCenterSecurityCmdlets | Export-Csv -NoTypeInformation "$filestamp-DataCenterSecurityCmdlet.csv"
$oDiscoverys | Export-Csv -NoTypeInformation "$filestamp-Discovery.csv"
$oExchangeAdmins | Export-Csv -NoTypeInformation "$filestamp-ExchangeAdmin.csv"
$oExchangeAggregatedOperations | Export-Csv -NoTypeInformation "$filestamp-ExchangeAggregatedOperation.csv"
$oExchangeItems | Export-Csv -NoTypeInformation "$filestamp-ExchangeItem.csv"
$oExchangeItemGroups | Export-Csv -NoTypeInformation "$filestamp-ExchangeItemGroup.csv"
$oMicrosoftTeamss | Export-Csv -NoTypeInformation "$filestamp-MicrosoftTeams.csv"
$oMicrosoftTeamsAddOnss | Export-Csv -NoTypeInformation "$filestamp-MicrosoftTeamsAddOns.csv"
$oMicrosoftTeamsSettingsOperations | Export-Csv -NoTypeInformation "$filestamp-MicrosoftTeamsSettingsOperation.csv"
$oOneDrives | Export-Csv -NoTypeInformation "$filestamp-OneDrive.csv"
$oPowerBIAudits | Export-Csv -NoTypeInformation "$filestamp-PowerBIAudit.csv"
$oSecurityComplianceCenterEOPCmdlets | Export-Csv -NoTypeInformation "$filestamp-SecurityComplianceCenterEOPCmdlet.csv"
$oSharePoints | Export-Csv -NoTypeInformation "$filestamp-SharePoint.csv"
$oSharePointFileOperations | Export-Csv -NoTypeInformation "$filestamp-SharePointFileOperation.csv"
$oSharePointSharingOperations | Export-Csv -NoTypeInformation "$filestamp-SharePointSharingOperation.csv"
$oSkypeForBusinessCmdletss | Export-Csv -NoTypeInformation "$filestamp-SkypeForBusinessCmdlets.csv"
$oSkypeForBusinessPSTNUsages | Export-Csv -NoTypeInformation "$filestamp-SkypeForBusinessPSTNUsage.csv"
$oSkypeForBusinessUsersBlockeds | Export-Csv -NoTypeInformation "$filestamp-SkypeForBusinessUsersBlocked.csv"
$oSways | Export-Csv -NoTypeInformation "$filestamp-Sway.csv"
$oThreatIntelligences | Export-Csv -NoTypeInformation "$filestamp-ThreatIntelligence.csv"
$oYammers | Export-Csv -NoTypeInformation "$filestamp-Yammer.csv"

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
