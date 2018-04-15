
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

$results = Search-UnifiedAuditLog -StartDate $start_date -EndDate $end_date -SessionCommand ReturnLargeSet -ResultSize 5000 #-RecordType AzureActiveDirectoryAccountLogon
$results | Out-File result.txt

$oUnifiedAuditLogs = @()
$oExchangeItems = @()
$oAzureActiveDirectoryStsLogons = @()
$oPowerBIAudits = @()
$oAzureActiveDirectoryAccountLogons = @()

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

    $oAuditData = convertfrom-json $result.AuditData

    
    if($result.RecordType -eq 'PowerBIAudit') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oItem += $oItem
    }

    if($result.RecordType -eq 'AzureActiveDirectoryAccountLogon') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oItem | Add-Member –MemberType NoteProperty –Name ExtendedPropertiesName –Value $result.AuditData.ExtendedProperties[0].Name
        $oItem | Add-Member –MemberType NoteProperty –Name ExtendedPropertiesValue –Value $result.AuditData.ExtendedProperties[0].Value 
        $oAzureActiveDirectoryAccountLogons += $oItem
    }

    if($result.RecordType -eq 'ExchangeItem') {
        $oItem = ConvertFrom-Json –InputObject $result.AuditData
        $oItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oExchangeItems += $oItem
    }


<#     if($result.RecordType -eq 'ExchangeItem') {
        $oExchangeItem = New-Object –TypeName PSObject
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name CreationTime –Value $oAuditData.CreationTime
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name Id –Value $oAuditData.Id
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name Operation –Value $oAuditData.Operation
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name OrganizationId –Value $oAuditData.OrganizationId
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name Workload –Value $oAuditData.Workload
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name UserId –Value $oAuditData.UserId
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name ClientIPAddress –Value $oAuditData.ClientIPAddress
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name ClientInfoString –Value $oAuditData.ClientInfoString
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name ExternalAccess –Value $oAuditData.ExternalAccess
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name InternalLogonType –Value $oAuditData.InternalLogonType
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name LogonType –Value $oAuditData.LogonType
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name LogonUserSid –Value $oAuditData.LogonUserSid
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name MailboxGuid –Value $oAuditData.MailboxGuid
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name MailboxOwnerSid –Value $oAuditData.MailboxOwnerSid
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name MailboxOwnerUPN –Value $oAuditData.MailboxOwnerUPN
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name OrganziationalName –Value $oAuditData.OrganziationalName
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name OriginatingServer –Value $oAuditData.OriginatingServer
        $oExchangeItems += $oExchangeItem
    }     #>

    if($result.RecordType -eq 'AzureActiveDirectoryStsLogon') {
        $oAzureActiveDirectoryStsLogon = New-Object –TypeName PSObject
        $oExchangeItem | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name CreationTime –Value $oAuditData.CreationTime
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name Id –Value $oAuditData.Id
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name Operation –Value $oAuditData.Operation
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name OrganizationId –Value $oAuditData.OrganizationId
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name RecordType –Value $oAuditData.RecordType
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name ResultStatus –Value $oAuditData.ResultStatus
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name UserKey –Value $oAuditData.UserKey
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name UserType –Value $oAuditData.UserType
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name Version –Value $oAuditData.Version
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name Workload –Value $oAuditData.Workload
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name ClientIP –Value $oAuditData.ClientIP
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name ObjectId –Value $oAuditData.ObjectId
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name UserId –Value $oAuditData.UserId
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name AzureActiveDirectoryEventType –Value $oAuditData.AzureActiveDirectoryEventType
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name ExtendedProperties –Value $oAuditData.ExtendedProperties
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name Actor –Value $oAuditData.Actor
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name ActorContextId –Value $oAuditData.ActorContextId
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name ActorIpAddress –Value $oAuditData.ActorIpAddress
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name InterSystemsId –Value $oAuditData.InterSystemsId
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name IntraSystemId –Value $oAuditData.IntraSystemId
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name Target –Value $oAuditData.Target
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name TargetContextId –Value $oAuditData.TargetContextId
        $oAzureActiveDirectoryStsLogon | Add-Member –MemberType NoteProperty –Name ApplicationId –Value $oAuditData.ApplicationId
        $oAzureActiveDirectoryStsLogons += $oAzureActiveDirectoryStsLogon
    }

<#     if($result.RecordType -eq 'PowerBIAudit') {
        $oPowerBIAudit = New-Object –TypeName PSObject
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name Id –Value $oAuditData.Id
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name RecordType –Value $oAuditData.RecordType
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name CreationTime –Value $oAuditData.CreationTime
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name Operation –Value $oAuditData.Operation
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name OrganizationId –Value $oAuditData.OrganizationId
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name UserKey –Value $oAuditData.UserKey
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name UserType –Value $oAuditData.UserType
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name Workload –Value $oAuditData.Workload
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name UserId –Value $oAuditData.UserId
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name ClientIP –Value $oAuditData.ClientIP
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name UserAgent –Value $oAuditData.UserAgent
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name Activity –Value $oAuditData.Activity
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name ItemName –Value $oAuditData.ItemName
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name WorkSpaceName –Value $oAuditData.WorkSpaceName
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name DatasetName –Value $oAuditData.DatasetName
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name ReportName –Value $oAuditData.ReportName
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name WorkspaceId –Value $oAuditData.WorkspaceId
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name ObjectId –Value $oAuditData.ObjectId
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name DatasetId –Value $oAuditData.DatasetId
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name ReportId –Value $oAuditData.ReportId
        $oPowerBIAudit | Add-Member –MemberType NoteProperty –Name IsSuccess –Value $oAuditData.IsSuccess
        $oPowerBIAudits += $oPowerBIAudit
    } #>


 <# 
    if($result.RecordType -eq 'AzureActiveDirectoryAccountLogon') {
       $oAzureActiveDirectoryAccountLogon = New-Object –TypeName PSObject
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name UnifiedAuditLogIdentity –Value $result.Identity
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name Id –Value $oAuditData.Id
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name CreationTime –Value $oAuditData.CreationTime
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name Operation –Value $oAuditData.Operation
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name OrganizationId –Value $oAuditData.OrganizationId
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name RecordType –Value $oAuditData.RecordType
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name ResultStatus –Value $oAuditData.ResultStatus
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name UserKey –Value $oAuditData.UserKey
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name UserType –Value $oAuditData.UserType
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name Version –Value $oAuditData.Version
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name Workload –Value $oAuditData.Workload
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name ClientIP –Value $oAuditData.ClientIP
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name ObjectId –Value $oAuditData.ObjectId
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name UserId –Value $oAuditData.UserId
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name AzureActiveDirectoryEventType –Value $oAuditData.AzureActiveDirectoryEventType
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name Client –Value $oAuditData.Client
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name LoginStatus –Value $oAuditData.LoginStatus
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name UserDomain –Value $oAuditData.UserDomain
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name ExtendedPropertiesName –Value $oAuditData.ExtendedProperties[0].Name
        $oAzureActiveDirectoryAccountLogon | Add-Member –MemberType NoteProperty –Name ExtendedPropertiesValue –Value $oAuditData.ExtendedProperties[0].Value 
        $oAzureActiveDirectoryAccountLogons += $oAzureActiveDirectoryAccountLogon
    }#>
}
$oUnifiedAuditLogs | Export-Csv -NoTypeInformation "oUnifiedAuditLogs-$filestamp.csv"
$oExchangeItems | Export-Csv -NoTypeInformation "oExchangeItems-$filestamp.csv"
$oAzureActiveDirectoryStsLogons | Export-Csv -NoTypeInformation "oAzureActiveDirectoryStsLogons-$filestamp.csv"
$oPowerBIAudits | Export-Csv -NoTypeInformation "oPowerBIAudits-$filestamp.csv"
$oAzureActiveDirectoryAccountLogons | Export-Csv -NoTypeInformation "oAzureActiveDirectoryAccountLogons-$filestamp.csv"

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
