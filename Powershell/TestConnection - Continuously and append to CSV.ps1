## This script will ping multiple IPs or Hostnames at the same time and log them with the date and time to a CSV file
## This will help in diagnosing if issues users are having are Internet issues, or Local network issues 
##e.g. If the ping to google.com fails, but to their default gateway does not, suggests internet

$csvPath = "" # Location to save the csv
$hosts = "8.8.8.8", "google.com", "bbc.co.uk", "10.207.136.1" # Hostnames or IPv4 adresses are valid

##################################################################################################################################

while ($true) { # Run Forver #
    $latencyResults = foreach ($hostname in $hosts) {

        $Connection = Test-Connection -Computername $hostname -Count 1 -ErrorAction SilentlyContinue

        $latencyResult = if ($Connection) { $Connection.ResponseTime } else { "Timeout" }

        [PSCustomObject]@{
            DateTime = Get-Date
            Hostname = $hostname
            LatencyOrTimeout = $latencyResult   
        }

        Write-Host "$hostname - $latencyResult" 
    }

    $latencyResults | Export-Csv -Path $csvPath -NoTypeInformation -Append # Adds to the existing spreadsheet

    Start-Sleep -Seconds 1  # Wait for a second before pinging again
}
## Written by Shunter 
## V1.0
## V1.1 > Cleaned up variables 
