# Exchange-2013-Database-and-Queue-

Gets Database and Queues statistics from Exchange servers.

Copy Powershell to C:\Program Files\BNGS.
Import Template to Zabbix.
Import Task Scheduler tasks to Exchnage Servers.

Add user account to authorize script execution, or allow local exzchange to have read permission on Exchange:
  C:\Users\dellserwis\Desktop>New-RoleGroup -Members exch01$,exch02$ -Name ËxchangeReadOnly -Roles "View-Only Configuration"
  
  
This project is using Send-ZabbixTrap:
https://www.powershellgallery.com/packages/Send-ZabbixTrap/1.2/Content/Send-ZabbixTrap.ps1
