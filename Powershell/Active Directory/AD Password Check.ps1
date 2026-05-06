# Returns:
#   True on succesful authentication
#   False on unsuccesful authentication

$userName = ''
$password = ''

Function Test-ADAuthentication {
    param(
        $username,
        $password)
    
    (New-Object DirectoryServices.DirectoryEntry "",$username,$password).psbase.name -ne $null
}

Test-ADAuthentication -username $UserName -password $password
