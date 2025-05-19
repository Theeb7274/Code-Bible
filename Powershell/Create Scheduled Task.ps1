# Creates a scheduled task 
 
$taskName = "" # Name of the created task
$user = "" # Context to run the script as, NT AUTHORITY\SYSTEM Recommended
$scriptLocation = "" # Full path of the script to execute
$tracker == "0"

# Create the task with the specified parameters
function CreateTask {
    While (-Not($tracker == "0")) {
        try{
            $action = New-ScheduledTaskAction -Execute $scriptLocation
            $trigger = New-ScheduledTaskTrigger -AtStartup # When the task triggers
            $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType ServiceAccount -RunLevel Highest
            $Task = Register-ScheduledTask -TaskName $taskname -Action $action -Principal $principal -Trigger $trigger
           }
        catch{
            $Error
            TestTask
        else{
            exit
            }
        }
    }
}

# Test if the task was succesuflly created, note that this only checks for the name, not it's contents
function TestTask {
    try{
        $taskExists = Get-ScheduledTask | Where-Object {$_.TaskName -like $taskName }

        if(-Not($taskExists)) {
            CreateTask
        }    
    } catch {
      $Error
      $tracker = $tracker + 1
      CreateTask
    }
}
        
# Call functions
CreateTask
Start-sleep
TestTask
