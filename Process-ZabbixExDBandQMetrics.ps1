$uri= "http://{0}/PowerShell/" -f ([System.Net.Dns]::GetHostByName(($env:computerName))).HostName

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos 
Import-PSSession $Session -CommandName get-queue,Get-MailboxDatabaseCopyStatus 


function Send-ZabbixTrap 
{
<#PSScriptInfo 
 
.VERSION 1.2 
 
.GUID 18bf0803-2626-46f5-b0d0-54e694985078 
 
.AUTHOR saw-friendship 
 
.COMPANYNAME 
 
.COPYRIGHT saw-friendship 
 
.TAGS Zabbix Sender OverTCP Json Send Trap 
 
.LICENSEURI 
 
.PROJECTURI https://sawfriendship.wordpress.com/ 
 
.ICONURI 
 
.EXTERNALMODULEDEPENDENCIES 
 
.REQUIREDSCRIPTS 
 
.EXTERNALSCRIPTDEPENDENCIES 
 
.RELEASENOTES 
 
 
#>

<# 
 
.DESCRIPTION 
 Send-ZabbixTrap send data to zabbix server over TCP without zabbix_sender.exe 
 
.LINK 
 https://sawfriendship.wordpress.com/ 
 
.EXAMPLE 
 Send-ZabbixTrap -z 172.16.5.2 -p 10051 -s Srv1 -k trap -o OK 
  
 You can use parameter aliases as in zabbix_sender.exe 
 z = zabbix server 
 p = port 
 s = host 
 k = key 
 o = value 
 
.EXAMPLE 
 Send-ZabbixTrap -z 172.16.5.2 -p 10051 -s Srv1 -k trap -o OK -OnlyPreview 
 
 Only generating Json string for preview 
  
.EXAMPLE 
 Send-ZabbixTrap -Server 172.16.5.2 -Port 10051 -JsonString $json 
  
 Json format example: 
  
 [pscustomobject][ordered]@{ 
 'request' = 'sender data'; 
 'data' = @( 
  1..3 | % { 
   [pscustomobject][ordered]@{ 
    'host' = 'HOST' 
    'key' = 'KEY' 
    'value' = 'VALUE' 
   } 
  } 
 ) 
} | ConvertTo-Json 
 
 
.EXAMPLE 
 Send-ZabbixTrap -Server 172.16.5.2 -Port 10051 -InputObject (Import-Csv -Encoding utf8 -Delimiter ';' -Path $home\srv.csv) -Header ComputerName,Service,Status 
 
 The PropertyNames of the input objects must be "host, key, value" or be specified in the header parameter 
 
#> 



[CmdletBinding(DefaultParameterSetName="Set0")]
param(
    [alias("z")]
    [string]$Server = '10.25.1.25',
    
    [alias("p")]
    [ValidateRange(1,65535)]
    [int]$Port = '10051',
    
    [parameter(ParameterSetName="Set1")]
    [alias("s")]
    [string]$HostName,
    
    [parameter(ParameterSetName="Set1")]
    [alias("k")]
    [string]$Key,
    
    [parameter(ParameterSetName="Set1")]
    [alias("o")]
    [string]$Value,
    
    [parameter(ParameterSetName="Set2")]$InputObject,
    
    [parameter(ParameterSetName="Set2",HelpMessage='Enter 3 string values for mapping object property to json headers. Default "host","key","value"')]
    [ValidateCount(3,3)]
    [string[]]$Header = @('host','key','value'),
    
    [parameter(ParameterSetName="Set3")]
    [string]$JsonString,
    
    [parameter(HelpMessage='On generating Json-string for preview')]
    [switch]$OnlyPreview
)

if ( [bool]($HostName -or $Key -or $Value) ) {
        if(! [bool]($HostName -and $Key -and $Value) ) {
            Write-Error 'HostName, Key and Value must not be null';
            break
        } else {
            $Json = [pscustomobject][ordered]@{
                'request' = 'sender data' ;
                'data' = @([pscustomobject][ordered]@{'host' = $HostName;'key' = $Key;'value' = $Value})
            } | ConvertTo-Json -Compress

    }
} elseif ($InputObject) {
    $Json = [pscustomobject][ordered]@{
                'request' = 'sender data' ;
                'data' = @(
                    $InputObject | Select-Object -Property @(
                        @{'Name' = 'host'; Expression = {$_.$($Header[0])}},
                        @{'Name' = 'key'; Expression = {$_.$($Header[1])}},
                        @{'Name' = 'value'; Expression = {$_.$($Header[2])}}
                    )
                )
            } | ConvertTo-Json -Compress
} elseif ($JsonString) {
    $Json = $JsonString | ConvertFrom-Json | ConvertTo-Json -Compress
} else {
    Write-Error 'Input data not found';
    break
}

if(!$Json){
    Write-Error 'Can not convert InputData to Json string';
    break
}

if($OnlyPreview){
    $Json | ConvertFrom-Json | ConvertTo-Json;
    break
}

try {
    [byte[]]$Header = @([System.Text.Encoding]::ASCII.GetBytes('ZBXD')) + [byte]1
    [byte[]]$Length = @([System.BitConverter]::GetBytes($([long]$Json.Length)))
    [byte[]]$Data = @([System.Text.Encoding]::ASCII.GetBytes($Json))
    
    $All = $Header + $Length + $Data
    
} catch {
    Write-Error 'Can not convert Json string to byte';
    break
}

try {
    $Socket = New-Object System.Net.Sockets.Socket ([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.SocketType]::Stream, [System.Net.Sockets.ProtocolType]::Tcp)
    $Socket.Connect($Server,$Port)
    $Socket.Send($All) | Out-Null
    [byte[]]$Buffer = New-Object System.Byte[] 1000
    [int]$ReceivedLength = $Socket.Receive($Buffer)
    $Socket.Close()
} catch {
    Write-Error 'TCP-level Error connecting, sending or receiving';
    break
}

$Received = [System.Text.Encoding]::ASCII.GetString(@($Buffer[13 .. ($ReceivedLength - 1)]))

try{
    $Received | ConvertFrom-Json
} catch {
    Write-Warning 'It is not possible to convert the output to a Json string, maybe the server has rejected invalid data'
    $Received
}
}

$queues=Get-Queue 

$queuesJson=""
$queuesJson+='{ "data": [ '
$queues | select -ExpandProperty NextHopDomain | %{$queuesJson+='{"{#EXQUEUE}": "'+$_+'" },'}
if ($queuesJson[$queuesJson.Length-1] -EQ ',') {$queuesJson=$queuesJson.Substring(0,$queuesJson.Length-1)}
$queuesJson+=']}'
$queuesJson
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key exqueue.discovery -Value $queuesJson

foreach ($queue in $queues)
{
$name=$queue.NextHopDomain
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exqueue.DeferredMessageCount[$name]" -Value $queue.DeferredMessageCount
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exqueue.LockedMessageCount[$name]" -Value $queue.LockedMessageCount
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exqueue.MessageCount[$name]" -Value $queue.MessageCount
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exqueue.RetryCount[$name]" -Value $queue.RetryCount
}

$databases = Get-MailboxDatabaseCopyStatus
$databasesJson=""
$databasesJson+='{ "data": [ '
$databases | select -ExpandProperty DatabaseName| %{$databasesJson+='{"{#EXDB}": "'+$_+'" },'}
if ($databasesJson[$databasesJson.Length-1] -EQ ',') {$databasesJson=$databasesJson.Substring(0,$databasesJson.Length-1)}
$databasesJson+=']}'
$databasesJson
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key exdb.discovery -Value $databasesJson

foreach ($database in $databases)
{
$name=$database.DatabaseName
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.ActiveCopy[$name]" -Value $database.ActiveCopy
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.ContentIndexBacklog[$name]" -Value $database.ContentIndexBacklog
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.ContentIndexRetryQueueSize[$name]" -Value $database.ContentIndexRetryQueueSize
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.ContentIndexState[$name]" -Value $database.ContentIndexState
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.Status[$name]" -Value $database.Status
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.IsLastCopyAvailabilityChecksPassed[$name]" -Value $database.IsLastCopyAvailabilityChecksPassed
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.IsLastCopyRedundancyChecksPassed[$name]" -Value $database.IsLastCopyRedundancyChecksPassed
if ($database.LatestFullBackupTime -ne $null) {Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.LatestFullBackupTime[$name]" -Value $database.LatestFullBackupTime}else{
                                               Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.LatestFullBackupTime[$name]" -Value "Never"}
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.LogCopyQueueIncreasing[$name]" -Value $database.LogCopyQueueIncreasing
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.LogReplayQueueIncreasing[$name]" -Value $database.LogReplayQueueIncreasing
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.MaxLogToReplay[$name]" -Value $database.MaxLogToReplay
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.ReplayQueueLength[$name]" -Value $database.ReplayQueueLength
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.CopyQueueLength[$name]" -Value $database.CopyQueueLength
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.ReplaySuspended[$name]" -Value $database.ReplaySuspended
Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.ReplicationIsInBlockMode[$name]" -Value $database.ReplicationIsInBlockMode
if ($database.LatestIncrementalBackupTime -ne $null) {Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.LatestIncrementalBackupTime[$name]" -Value $database.LatestIncrementalBackupTime}else{
                                                      Send-ZabbixTrap -HostName $env:COMPUTERNAME.ToLower() -Key "exdb.LatestIncrementalBackupTime[$name]" -Value "Never"}

}
Remove-PSSession $Session 
