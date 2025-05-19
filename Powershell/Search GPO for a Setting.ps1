# Searches for the GPO Specified when prompted, edit the scope as required
 
# Get the string we want to search for 
$string = Read-Host -Prompt "GPO Setting you want to search for?" 
 
# Set the domain to search for GPOs 
$DomainName = $env:USERDNSDOMAIN 
 
# Find all GPOs in the current domain 
write-host "Finding all the GPOs in $DomainName" 
Import-Module grouppolicy 

$allGposInDomain = Get-GPO -All -Domain $DomainName 
[string[]] $MatchedGPOList = @()

# Look through each GPO's XML for the string 
Write-Host "Starting search...." 

foreach ($gpo in $allGposInDomain) { 
    $report = Get-GPOReport -Guid $gpo.Id -ReportType Xml 
    if ($report -match $string) { 
        write-host "********** Match found in: $($gpo.DisplayName) **********" -foregroundcolor "Green"
        $MatchedGPOList += "$($gpo.DisplayName)";
    } 

    else { 
        Write-Host "No match in: $($gpo.DisplayName)" 
    }
} 
write-host "`r`n"
write-host "Results: **************" -foregroundcolor "Yellow"

foreach ($match in $MatchedGPOList) { 
    write-host "Match found in: $($match)" -foregroundcolor "Green"
}

Read-Host "Press Enter to close"
