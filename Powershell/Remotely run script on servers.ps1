$Credential = Get-Credential
$Array = Invoke-Command -ComputerName (Get-Content C:\Users\Admin\Desktop\Shtuff\Computers_no_domain.txt) -Credential $Credential -ScriptBlock { Get-ScheduledTask | Where-Object {$_.TaskName -like $taskName } }
write-host $Array 

#$RemoteSession = New-PSSession -ComputerName $RemoteComputer -Credential $Credential
#Invoke-Command -Session $RemoteSession -ScriptBlock {
#   Get-Host
#}