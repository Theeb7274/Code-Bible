function ServiceCheck {

    $serviceName = "Datto RMM"
    $status = (Get-Service -Name $serviceName).status
        if ($status -ne "Running") {
        Start-Service -Name $serviceName
    }
}

ServiceCheck