# How to use
## Create Module folder
`%USERPROFILE%\Documents\WindowsPowerShell\Modules\Monitor-Mail`  

## Save file in module folder  
`%USERPROFILE%\Documents\WindowsPowerShell\Modules\Monitor-Mail\Monitor-Mail.psm1`  

## Open Powershell and run following command  
`$ Import-Module -Name Monitor-Mail -Verbose`  

## Schedule a task in Task Scheduler to run following command  
`Monitor-Mail -FolderName "Inbox" -MinAgo 70 -MailSubject "Mail to watch"`
https://marte-it.at/en/start-powershell-script-hidden-via-task-scheduler/
