$script = "\\ad.sfet.org.uk\SysVol\ad.sfet.org.uk\Policies\{8EAA54E3-ABDB-4081-9659-65F22B191C12}\Machine\Scripts\Shutdown\del_prof_wrapper.ps1"
$exclusions = "admin", "administrator", "public", "default"
$path = "\\sfet-bhcs-mgt01\software$\Delprof2 1.6.0"
$logfile = "C:\EDUIT\Delprof_$(Get-Date -F 'yyyyMMdd-HHmmss')log.txt"

& $script -Unattended -ExcludeUsers $exclusions -Verbose -Logfile $logfile -DelprofPath $path -ListOnly