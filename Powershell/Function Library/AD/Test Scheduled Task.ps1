# Tests that a scheduled task exists, it does not check any actions of the task
function TestTask {
    try{
        $taskName = ""
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