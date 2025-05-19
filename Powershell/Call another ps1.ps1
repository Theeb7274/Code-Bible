# Calls another ps1 and passes paramaters onto it
# If the script allows it, this can be used to pass general parameters aswell, such as -verbose
# & tells powershell the following text is a script to be executed, otherwise it doesn't handle it correctly

# Configuration
$script = ""

# Call $script with parameters
& $script 
