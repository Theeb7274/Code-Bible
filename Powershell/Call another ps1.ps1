# Calls another ps1 and passes paramaters onto it
# If the script allows it, this can be used to pass general parameters aswell, such as -verbose
# & is a call operator which makes sure powershell handles $script as a command to be executed

# Configuration
$script = ""

# Call $script with parameters
& $script 
