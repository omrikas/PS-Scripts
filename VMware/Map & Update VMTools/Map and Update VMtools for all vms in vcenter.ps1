Import-Module "VMware.VimAutomation.Core"
Connect-VIServer (read-host "Enter vCenter: ")

# Define arrays
[array]$NotInstalled = @{}
[array]$NotUpdated = @{}
[array]$NotRunning = @{}
[array]$UpToDate = @{}
[array]$Catcher = @{}

# Get vms
$vms = get-vm | ? {($_.PowerState.ToString().Contains("On")) -and ($_.Guest.OSFullName -ne $null) -and ($_.Guest.OSFullName.tolower().Contains("windows"))} 

# Check all vms
foreach ($vm in $vms)
{
    $vmview = get-view -VIObject $vm
    # Act according to the tool status
    switch ($vmview.Guest.ToolsStatus)
    {
        'toolsNotInstalled' 
        {
            $NotInstalled += $vm.name
        }
        'toolsOld' 
        {
            $NotUpdated += $vm.name
  
            # Get free space on C:\ drive
            $LocalDisk = $vm.Guest.disks | ? {($_.path.contains("C")) -or ($_.path -eq '/')}
            $FreeSpaceGB = [math]::Round(($LocalDisk.FreeSpace / 1gb) , 2)
            
            # Check if the space is greater than 1gb
            if ($FreeSpaceGB -gt 1)
            {
                update-tools -VM $vm -NoReboot -RunAsync
            }
            else
            {
                "Drive size $([math]::Round($localdisk.CapacityGB,0))gb $($vm.name) "
            }
        }
        'toolsNotRunning' 
        {
            $NotRunning += $vm.name
        }
        'toolsOk' 
        {
            $UpToDate += $vm.name
            
        }
        default
        {
            $Catcher += $vm.name            
        }
    }
}

Disconnect-VIServer * -Confirm:$false