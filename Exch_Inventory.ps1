$Now = Get-Date
$Report_Path = Split-Path -Parent $MyInvocation.MyCommand.Path
$Report_File = $Report_Path+"\Exchange_Inventory_Report.txt"
$Detail_File = $Report_Path+"\Exchange_Detailed_Report.txt"

$ProgressActivity = "Initializing..."
$msgString = "Initializing Collection"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 0
Write-Verbose $msgString

# Write Report information
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Exchange Inventory Report" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
Write-Output "  Report Date: $Now " >> $Report_File
Write-Output "" >> $Report_File
Write-Output "" >> $Report_File

$ProgressActivity = "Collecting..."

$sysinfo = Get-WmiObject -Class Win32_ComputerSystem
$hostname = $sysinfo.Name+"."+$sysinfo.Domain
$hosttype = $sysinfo.model
Write-Output "  Report running on $hostname " >> $Report_File
Write-Output "  The system is $hosttype" >> $Report_File
Write-Output "  Domain Controller list: " >> $Report_File
Get-DomainController | Format-Table dnshostname,Adsite >> $Report_File
Write-Output ""

$msgString = "Collecting data about Event Log"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 1
Write-Verbose $msgString
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* EventLog Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
Get-EventLog -LogName Application -Newest 300 | Where-Object {$_.entrytype -eq ('Warning' -or 'Error' -or 'Critical')} | Group-Object -Property entrytype >> $Report_File
Get-EventLog -LogName Application -Newest 1000 | Where-Object {$_.entrytype -eq ('Warning' -or 'Error' -or 'Critical')} | Select-Object EntryType,EventID,Source,Category,Message >> $Detail_File

# Get Mailbox information
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Mailbox Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
$msgString = "Collecting data about Mailbox"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 3
Write-Verbose $msgString
Write-Output "Total Mailbox: $((Get-Mailbox -ResultSize unlimited).count)" >> $Report_File
Write-Output "Resource Mailbox: $((Get-mailbox -resultsize unlimited | Where-Object {$_.isresource -eq $True}).count)" >> $Report_File
Write-Output "Shared Mailbox: $((Get-mailbox -resultsize unlimited | Where-Object {$_.isshared -eq $True}).count)" >> $Report_File
Write-Output "Linked Mailbox: $((Get-mailbox -resultsize unlimited | Where-Object {$_.islinked -eq $True}).count)" >> $Report_File
Write-Output "" >> $Report_File

$msgString = "Collecting data about Mailbox"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 5
Write-Verbose $msgString

Write-Output "Count by Server:" >> $Report_File
Get-Mailbox | Group-Object -Property:ServerName | Select-Object Name,Count >> $Report_File
Write-Output "Count by Database:" >> $Report_File
Get-Mailbox | Group-Object -Property:Database | Select-Object Name,Count >> $Report_File
Write-Output "Total Recipient: $((Get-Recipient -resultsize unlimited).count)" >> $Report_File
Write-Output "" >> $Report_File
Write-Output "Count by Recipient type: " >> $Report_File
Get-Recipient -resultsize unlimited | Group-Object -Property:RecipientType | Select-Object Name,Count >> $Report_File
Write-Output "Total Contact: $((Get-Recipient -resultsize unlimited).count)" >> $Report_File
Write-Output "" >> $Report_File
Write-Output "Count by Contact type: " >> $Report_File
Get-Recipient -resultsize unlimited | Group-Object -Property:RecipientType | Select-Object Name,Count >> $Report_File

$msgString = "Collecting data about Mailbox"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 9
Write-Verbose $msgString

# Get Management Roles informatin
$msgString = "Collecting data about Management Roles"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 11
Write-Verbose $msgString

Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Management Roles Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
Get-RoleGroupMember -Identity "Organization Management" | Select-Object Name,OrganizationalUnit >> $Report_File
Get-RoleGroupMember -Identity "View-Only Organization Management" | Select-Object Name,OrganizationalUnit >> $Report_File

# Get Address List information
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Address List Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
$msgString = "Collecting data about Address List"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 13
Write-Verbose $msgString

Get-AddressList | Format-List >> $Detail_File
Write-Output "Address List: " >> $Report_File
Get-AddressList >> $Report_File

Get-GlobalAddressList | Format-List >> $Detail_File
Write-Output "Global Address List: " >> $Report_File
Get-GlobalAddressList | Format-Table >> $Report_File


# Get Accept Domains
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Accepted Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
$msgString = "Collecting data about Accepted Domains"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 15
Write-Verbose $msgString
Write-Output "Accepted Domains: " >> $Report_File
Get-AcceptedDomain >> $Report_File

# Get E-mail Address Policy information
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* E-mail Address Policy Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
$msgString = "Collecting data about E-Mail Address Policies"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 17
Write-Verbose $msgString
Get-EmailAddressPolicy | Format-List >> $Detail_File
Write-Output "E-Mail Address Policies: " >> $Report_File
Get-EmailAddressPolicy | Select-Object Name,Priority,RecipientFilter,LdapRecipientFilter,EnabledPrimarySMTPAddressTemplate,EnabledEmailAddressTemplates >> $Report_File

# Get Receive Connector
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Receive Connector Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
$msgString = "Collecting data about Receive Connector"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 19
Write-Verbose $msgString
Get-ReceiveConnector | Format-List >> $Detail_File
Write-Output "Receive Connectors: " >> $Report_File
Get-ReceiveConnector >> $Report_File

# Get Send Connector
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Send Connector Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
$msgString = "Collecting data about Send Connector"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 21
Write-Verbose $msgString
Get-SendConnector | Format-List >> $Detail_File
Write-Output "Send Connectors: " >> $Report_File
Get-SendConnector >> $Report_File

# Get Public Folder/Mailbox information

# Get Exchange Server infomation
Write-Output "" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Exchange Server Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
$msgString = "Collecting data about Exchange Server"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 23
Write-Verbose $msgString
Get-ExchangeServer -Status | Format-List >> $Detail_File
Write-Output "Total Exchange Servers: $((Get-ExchangeServer).count)" >> $Report_File
Write-Output "" >> $Report_File
$msgString = "Collecting data about Exchange Server"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 25
Write-Verbose $msgString
Write-Output "Exchange Servers List:" >> $Report_File
Get-ExchangeServer | Select-Object Name,ServerRole,Edition,ExchangeVersion,Domain,DataPath >> $Report_File
$msgString = "Collecting data about Exchange Server"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 27
Write-Verbose $msgString
Write-Output "Group By Site: " >> $Report_File
Get-ExchangeServer | Group-Object -property:Site | Format-List Name,Count,Group >> $Report_File
$msgString = "Collecting data about Exchange Server"
Write-Output "" >> $Report_File
Write-Output "Group By Role: " >> $Report_File
$msgString = "Collecting data about Exchange Server"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 29
Write-Verbose $msgString
Write-Output " Client Access Server: " >> $Report_File
Get-ExchangeServer | Where-Object {$_.ServerRole -like "*ClientAccess*"} | Group-Object -property:ServerRole | Format-List Name,Count,Group >> $Report_File
$msgString = "Collecting data about Exchange Server"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 31
Write-Verbose $msgString
Write-Output " Hub Transport Server: " >> $Report_File
Get-ExchangeServer | Where-Object {$_.ServerRole -like "*HubTransport*"} | Group-Object -property:ServerRole | Format-List Name,Count,Group >> $Report_File
$msgString = "Collecting data about Exchange Server"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 33
Write-Verbose $msgString
Write-Output " Mailbox Server: " >> $Report_File
Get-ExchangeServer | Where-Object {$_.ServerRole -like '*Mailbox*'} | Group-Object -property:ServerRole  | Format-List Name,Count,Group >> $Report_File
$msgString = "Collecting data about Exchange Server"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 35
Write-Verbose $msgString
Write-Output " Edge Server: " >> $Report_File 
Get-ExchangeServer | Where-Object {$_.ServerRole -like '*Edge*'} | Group-Object -property:ServerRole | Format-List Name,Count,Group >> $Report_File
Write-Output "" >> $Report_File

# Get Mailbox database information
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Mailbox Database Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
$msgString = "Collecting data about Mailbox Database"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 35
Write-Verbose $msgString
Get-MailboxDatabase -Status | Format-List >> $Detail_File
Write-Output "Mailbox Database Size:" >> $Report_File
Get-MailboxDatabase -Status | Select-Object Name,ServerName,DatabaseSize >> $Report_File

# Get Database Availability Group information
$msgString = "Collecting data about DAG"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 37
Write-Verbose $msgString
Get-DatabaseAvailabilityGroup | Format-List >> $Detail_File

# Get virtual directories information
Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Virtual Directories Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
$msgString = "Collecting data about Virtual Directories"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 39
Write-Verbose $msgString
Get-AutodiscoverVirtualDirectory | Format-List >> $Detail_File
Write-Output "Autodiscover virtual directory by version group:" >> $Report_File
Get-AutodiscoverVirtualDirectory | Group-Object -Property:AdminDisplayVersion >> $Report_File
Write-Output "Autodiscover virtual directory List:" >> $Report_File
Get-AutodiscoverVirtualDirectory | Select-Object Name,ServerName,AdminDisplayVersion,InetrnalURL,ExternalURL >> $Report_File

$msgString = "Collecting data about Virtual Directories"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 46
Write-Verbose $msgString
Get-ECPVirtualDirectory | Format-List >> $Detail_File
Write-Output "ECP virtual directory by version group:" >> $Report_File
Get-ECPVirtualDirectory | Group-Object -Property:AdminDisplayVersion >> $Report_File
Write-Output "ECP virtual directory List:" >> $Report_File
Get-ECPVirtualDirectory | Select-Object Name,Server,AdminDisplayVersion,InetrnalURL,ExternalURL >> $Report_File

$msgString = "Collecting data about Virtual Directories"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 53
Write-Verbose $msgString
Get-WebServicesVirtualDirectory | Format-List >> $Detail_File
Write-Output "WebServices virtual directory by version group:" >> $Report_File
Get-WebServicesVirtualDirectory | Group-Object -Property:AdminDisplayVersion >> $Report_File
Write-Output "WebServices virtual directory List:" >> $Report_File
Get-WebServicesVirtualDirectory | Select-Object Name,Server,AdminDisplayVersion,InetrnalURL,ExternalURL >> $Report_File

$msgString = "Collecting data about Virtual Directories"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 60
Write-Verbose $msgString
Get-ActiveSyncVirtualDirectory | Format-List >> $Detail_File
Write-Output "ActiveSync virtual directory by version group:" >> $Report_File
Get-ActiveSyncVirtualDirectory | Group-Object -Property:AdminDisplayVersion >> $Report_File
Write-Output "ActiveSync virtual directory List:" >> $Report_File
Get-ActiveSyncVirtualDirectory | Select-Object Name,Server,AdminDisplayVersion,InetrnalURL,ExternalURL >> $Report_File

$msgString = "Collecting data about Virtual Directories"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 67
Write-Verbose $msgString
Get-OABVirtualDirectory | Format-List >> $Detail_File
Write-Output "OAB virtual directory by version group:" >> $Report_File
Get-OABVirtualDirectory | Group-Object -Property:AdminDisplayVersion >> $Report_File
Write-Output "OAB virtual directory List:" >> $Report_File
Get-OABVirtualDirectory | Select-Object Name,Server,AdminDisplayVersion,InetrnalURL,ExternalURL >> $Report_File

$msgString = "Collecting data about Virtual Directories"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 74
Write-Verbose $msgString
Get-OWAVirtualDirectory | Format-List >> $Detail_File
Write-Output "OWA virtual directory by version gourp:" >> $Report_File
Get-OWAVirtualDirectory | Group-Object -Property:AdminDisplayVersion >> $Report_File
Write-Output "OWA virtual directory List:" >> $Report_File
Get-OWAVirtualDirectory | Select-Object Name,ServerName,AdminDisplayVersion,InetrnalURL,ExternalURL,OWAVersion >> $Report_File

$msgString = "Collecting data about Virtual Directories"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 81
Write-Verbose $msgString
Get-PowershellVirtualDirectory | Format-List >> $Detail_File
Write-Output "Powershell virtual directory by version group:" >> $Report_File
Get-PowershellVirtualDirectory | Group-Object -Property:AdminDisplayVersion >> $Report_File
Write-Output "Powershell virtual directory List:" >> $Report_File
Get-PowershellVirtualDirectory | Select-Object Name,Server,AdminDisplayVersion,InetrnalURL,ExternalURL >> $Report_File

$msgString = "Collecting data about Virtual Directories"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 88
Write-Verbose $msgString
Get-MAPIVirtualDirectory | Format-List >> $Detail_File
Write-Output "MAPI virtual directory by version group:" >> $Report_File
Get-MAPIVirtualDirectory | Group-Object -Property:AdminDisplayVersion >> $Report_File
Write-Output "MAPI virtual directory List:" >> $Report_File
Get-MAPIVirtualDirectory | Select-Object Name,Server,AdminDisplayVersion,InetrnalURL,ExternalURL >> $Report_File

# Get Certificates information
$msgString = "Collecting data about Certificates"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 95
Write-Verbose $msgString

Write-Output "*****************************************************************" >> $Report_File
Write-Output "* Certificates Information" >> $Report_File
Write-Output "*****************************************************************" >> $Report_File
Get-ExchangeCertificate | Format-List >> $Detail_File
Write-Output "Imported certificates:" >> $Report_File
Get-ExchangeCertificate | Select-Object Subject,CertificateDomains,Services,NotBefore,NotAfter,IsSelfSigned >> $Report_File

$msgString = "Finished"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 100
Write-Verbose $msgString
