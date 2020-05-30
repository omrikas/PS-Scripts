import-module VMware.VimAutomation.Core

#Inputs
[string]$vcenter = read-host "Please enter vcenter name"
[string]$folder = read-host "Please enter the folder name of the vms you want to map"
[string]$netapp = read-host "Please enter netapp name"

<# If you wish to have a hard coded user to login, use this connection and fill the usr/pass
$usr = ""
$pass = ""
$encryptedPass = New-Object -TypeName System.Security.SecureString
$pass.ToCharArray() | % {$encryptedPass.AppendChar($_)}
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $usr,$encryptedPass
Connect-VIServer $vcenter -Credential $cred
#>

# Connect to vcenter
"Connecting to vCenter"
Connect-VIServer $vcenter 

<# If you wish to have a hard coded user to login, use this connection and fill the usr/pass
$usr = ""
$pass = ""
$encryptedPass = New-Object -TypeName System.Security.SecureString
$pass.ToCharArray() | % {$encryptedPass.AppendChar($_)}
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $usr,$encryptedPass
Connect-NcController $netapp -Credential $cred -HTTPS
#>

# Connect to Netapp
"Connecting to Netapp"
Connect-NcController $netapp -HTTPS

# Get Vols and Luns
"Fetching luns and vols"
$SRCVols = Get-NcVol
$SRCLuns = Get-NcLunMap
$MirrorDestinations = Get-NcSnapmirrorDestination 

# Get VMs with RDMs
"Getting VMs with RDMS"
$FolderName = Read-Host "Enter Folder Name: "
$VMsWithRDMs = Get-Folder -name "$FolderName" | Get-VM | ? {(Get-HardDisk $_ -DiskType RawPhysical) -ne $null} | Sort-Object
$VMsTable = @()

# Map all vms with RDMs
"Mapping VMs with RDMs"
foreach ($currVM in $VMsWithRDMs)
{
    # Get all RDMs
    $RDMs = get-harddisk $currVM -DiskType RawPhysical
    $LunIdList = @()

    # Map each RDM to the lun id
    foreach ($CurrRDM in $RDMs)
    {
        $esxcli = Get-EsxCli -VMHost $currVM.VMHost -v2
        $LunIdList += $esxcli.storage.nmp.path.list.Invoke(@{'device'=$CurrRDM.ScsiCanonicalName}).RuntimeName.Split(':')[-1].TrimStart('L')
    }

    $VMsTable += $currVM | select @{name="VM"; Expression={$currVM.Name}}, 
                                  @{name="Cluster"; Expression={$currVM.VMHost.Parent.name.ToLower()}}, 
                                  @{Name="LunIDs"; Expression={$LunIdList}}
}

# Check each vm in the table
"Checking For Snapmirrors"
$ResultTable = @()
foreach ($currRow in $VMsTable)
{
    # Find the matching volumes for each of the vm luns
    foreach ($currLunID in $currRow.LunIDs)
    {
        $VolName = ($SRCLuns | ? {($_.InitiatorGroup -eq $currRow.Cluster) -and ($_.LunId -eq $currLunID)}).Path.ToString().Split("/")[2]
        $CurrVol = $SRCVols | ? {$_.name -eq $volname}
        $hasMirror = "No"
        if ($MirrorDestinations | ? {$_.SourceVolume -eq $currvol.name})
        {
            $hasMirror = "Yes"
        }

        $ResultTable += $CurrVol | select @{name="VM"; Expression={$currRow.VM}}, 
                                          @{name="Cluster"; Expression={$currRow.Cluster}},
                                          @{Name="Vol Name"; Expression={$VolName}},
                                          @{Name="LunIDs"; Expression={$currLunID}},
                                          @{Name="Has Mirror"; Expression={$hasMirror}}

    }
}

Disconnect-VIServer * -Force:$true -Confirm:$false
$resulttable | Export-Csv -LiteralPath c:\temp\SnapmirrorCheck\$CurrSite.csv -Force:$true