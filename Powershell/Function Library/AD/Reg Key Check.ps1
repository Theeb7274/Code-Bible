# Checks for the given key, returns a boolean
function CheckRegKeyExists {
    param (
            [parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]$Path,

            [parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]$Value
          )

    try {
        Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
        #Write-Host "Key Exists"
        return $true         
    }
     catch 
     {
        #Write-Host "Key Does not Exist"
        return $false
    }                        
}