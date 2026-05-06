# Adding $group to users, pulled from a csv
# Lightweight design by avoiding module imports & unecssary functions
# No modules imported, given how likely an DC already has the module. AD module required.

# Variables

$CsvPath = ""
$group = ""
$header = "" # csv header the for loop should use

# Function 
function Add-Group {
    Add-ADgroupMember -Identity $group -Members $user.$header
}

# Main loop imports the whole csv and loops through the contents of $header
$users = Import-Csv $CsvPath

foreach ($user in $users) {
    try {
        Add-AdGroupMember
    }
    catch {
        Write-Host $_Exception.Message
    }
}
