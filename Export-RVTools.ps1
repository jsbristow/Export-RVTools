# Export RVTools reports to .xlsx files for multiple vCenters and/or unmanaged ESXi hosts
# Requires modules:
#   VMware.VimAutomation.Core
#   Microsoft.PowerShell.SecretManagement
#   Microsoft.PowerShell.SecretStore
# 2022-12-10 - JSB

# Import and filter list of vCenters and unmanaged ESXi hosts from CSV or change to your method. Field name 'Name' used.
$vCenters = Import-Csv "E:\vmScripts\vCenters.csv" | where {$_.Hypervisor -match "VMware vCenter"}
$vHosts = Import-Csv "E:\vmScripts\vHosts.csv" | where {$_.HyperVisor -match "VMware ESXi" -and $_.vCenter -eq ""}

$vCenters = $vCenters.Name
$vHosts   = $vHosts.Name

$date = (Get-Date -Format "yyyy-MM-dd")
$year = (Get-Date -Format "yyyy")

$ExportPath = "E:\vmScripts\Export-RVTools\$year\$date"

# Credentials - Use 'C:\Program Files (x86)\Robware\RVTools\RVToolsPasswordEncryption.exe' to encrypt password for RVTools.
# vCenter credential in SPN format stored in SecretStore.
$userid = "vcuser@example.com"
$password = (Get-Secret 'RVTools_vcuser').GetNetworkCredential().Password
# ESXi credential
$userid2 = "root"
$password2 = (Get-Secret 'RVTools_root').GetNetworkCredential().Password

# Creates the Folder & Path with current date
New-Item -Path $ExportPath -ItemType Directory -ErrorAction SilentlyContinue

# Export RVTools report for each vCenter Server
$i = 0
foreach ($vCenter in $vCenters){
    # Run Export
    Write-Progress -Activity "RVTools Exporting: $vCenter" -PercentComplete (($i / $vCenters.Count) * 100)
    $Arguments = "-passthroughAuth -s $vCenter -c ExportAll2xlsx -d $ExportPath -f $vCenter.xlsx"
    Start-Process -FilePath "C:\Program Files (x86)\RobWare\RVTools\RVTools.exe" -ArgumentList $Arguments -PassThru -Wait
    # . "C:\Program Files (x86)\RobWare\RVTools\RVTools.exe" -passthroughAuth -s $vCenter -c ExportAll2xlsx -d $ExportPath -f "$vCenter.xlsx"
    $i++
}

# Export RVTools report for each unmanaged ESXi host
$i = 0
foreach ($vHost in $vHosts){
    # Run Export
    Write-Progress -Activity "RVTools Exporting: $vHost" -PercentComplete ($i / $vHosts.Count * 100)
    $Arguments = "-u $userid2 -p $password2 -s $vHost -c ExportAll2xlsx -d $ExportPath -f $vHost.xlsx"
    Start-Process -FilePath "C:\Program Files (x86)\RobWare\RVTools\RVTools.exe" -ArgumentList $Arguments -PassThru -Wait
    # . "C:\Program Files (x86)\RobWare\RVTools\RVTools.exe" -u $userid2 -p $password2 -s $vHost -c ExportAll2xlsx -d $ExportPath -f "$vHost.xlsx"
    $i++  
}

# End