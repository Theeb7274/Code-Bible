# Create the task with the specified parameters
function CreateTask {
    $taskName = "" # Name of the created task
    $user = "" # Context to run the script as, NT AUTHORITY\SYSTEM Recommended
    $scriptLocation = "" # Full path of the script to execute
    $tracker == "0"
    
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