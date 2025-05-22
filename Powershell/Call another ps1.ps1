# Calls another ps1 and passes paramaters onto it
# If the script allows it, this can be used to pass general parameters aswell, such as -verbose
# & is a call operator which makes sure powershell handles $script as a command to be executed

# Configuration
$script = "" # Either include the full path to the ps1, or change the working directory to where the script is and use .\script_name_here
# Modify these as per your script requires
$Parameter1 = ""
$Parameter2 = ""
$Parameter3 = ""

# Call $script with parameters
# An example would be & $script -GroupName $Parameter1 -HostName $Parameter 2 -verbose
& $script
