# Starts $serviceName if it's stopped
# This script isn't a reliable way of keeping a service alive, this was made to catch a remote access service which has been flakey
# This will be revisited if a more robust one is required

# Function to check the status of $serviceName and start if stopped 
 function ServiceCheck {

    $serviceName = ""
    $status = (Get-Service -Name $serviceName).status
        if ($status -ne "Running") {
        Start-Service -Name $serviceName
    }
}

# Call function
ServiceCheck
