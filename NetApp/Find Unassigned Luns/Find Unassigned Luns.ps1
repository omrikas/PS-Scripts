import-module "VMware.VimAutomation.Core"
Connect-VIServer (read-host "Enter vCenter")
Connect-NcController (read-host "Enter netapp") -https
$clusters = Get-Cluster
$luns = get-nclun

# Check all clusters
foreach ($cluster in $clusters)
{
    $scsiTab = @{}
    $UnmappedLuns = @{}

    # Run the check on one esxi from the cluster (takes way too long to be ran on all esxis)
    $esxi = get-cluster $cluster | Get-VMHost | ? {$_.connectionstate -ne "Maintenance"} | Select-Object -First 1
    $esxi | get-scsilun | where {$_.luntype -eq "Disk"} | % {
                                                                 $key = $cluster.name + "-" + $_.canonicalname.split(".")[1]
                                                                 if (!$scsiTab.ContainsKey($key))
                                                                 {
                                                                     $scsitab[$key] = $_.canonicalname,"", (Get-ScsiLunPath -ScsiLun $_ | ? {$_.state -eq "Active"} | Select-Object -First 1).lunpath.split("L")[1], [math]::Round($_.capacitygb, 0), $cluster.name
                                                                 }
                                                            }

    # Add all vms to the scsi tab
    $vms = get-vm | ? {$_.VMHost.parent -eq $cluster} | Sort-Object
    foreach ($vm in $vms)
    {
        $Rdms = Get-HardDisk $vm | ? {$_.Persistence.ToString().contains("Independent")}
        foreach ($Rdm in $rdms)
        {
                $key = $cluster.name.Split(".")[0] + "-" + $rdm.ScsiCanonicalName.split(".")[1]
                if ($scsiTab[$key] -ne $null)
                {
                    $scsitab[$key][1] = $vm.name
                }
        }
    }

    # Add all datastores to the scsi tab
    $datastores = Get-Datastore
    foreach ($datastore in $datastores)
    {
            $key = $cluster.name.Split(".")[0] + "-" + $datastore.ExtensionData.Info.Vmfs.Extent[0].DiskName.split(".")[1]
            if ($scsitab[$key] -ne $null)
            {
                $scsitab[$key][1] = $datastore.name
            }
    }

    # Get all unmapped luns
    foreach ($ScsiName in $scsiTab.Keys)
    {
        if ($scsiTab[$ScsiName][1] -eq "")
        {
            $UnmappedLuns[$ScsiName] += $scsiTab[$ScsiName]
        }    
    }

    # Map all unmapped luns to the corresponding volume
    foreach ($lun in $luns)
    { 
        $lunmap = $null
        foreach ($UnmappedLun in $UnmappedLuns.Keys)
        {
            # Check if its the same size
            if (([math]::Round($lun.Size / (1GB), 0)) -eq $UnmappedLuns[$UnmappedLun][3])
            {
                $LunMap = Get-NcLunMap $lun  

                # Check if its the same lun id and mapped to the same igroup
                if (($lunmap.LunId -eq $UnmappedLuns[$UnmappedLun][2]) -and ($lunmap.InitiatorGroup.tolower() -eq $UnmappedLuns[$UnmappedLun][4].ToLower()))
                {
                    $UnmappedLuns[$UnmappedLun][1] += $lun.Volume + " "
                }
            }
        }
    }

    [array]$FinalVms = $null
    foreach ($UnmappedLun in $UnmappedLuns.Keys)
    {
        $FinalVM = New-Object object
        $FinalVM | Add-Member -MemberType NoteProperty -Name Volume -Value $UnmappedLuns[$UnmappedLun][1]
        $FinalVM | Add-Member -MemberType NoteProperty -Name Size -Value $UnmappedLuns[$UnmappedLun][3]
        $FinalVM | Add-Member -MemberType NoteProperty -Name LunID -Value $UnmappedLuns[$UnmappedLun][2]

        $FinalVMS += $FinalVM 
        $Path = "c:\temps\" + $cluster + ".csv"
        $finalvms | Export-Csv $Path
    }
}