#################################################################################################################
#### This script checks all non mirror volumes on the san vserver, deletes old snapshots and resizes volumes ####
#################################################################################################################

# Deletes all old snapshots on volume
function DeleteOldSnapshots($volume)
{
    # Get all snapshots on vol
    $Snapshots = Get-NcSnapshot -Volume $volume | ? {($_.dependency -eq $null)}
    
    # Check each snapshot to see if it should be deleted
    foreach ($snapshot in $Snapshots)
    {
        # Delete old snapshots
        if ($Snapshot.AccessTimeDT.AddDays(30) -lt (get-date))
        {
            $snapshot.name + " " + ((get-date) - $snapshot.AccessTimeDT).days
            Add-Content $log "Deleting $snapshot, Age:  $(((get-date) - $snapshot.AccessTimeDT).days) Days"
            Remove-NcSnapshot -Snapshot $snapshot -Volume $snapshot.Volume -VserverContext $snapshot.Vserver -Confirm:$false
        }
    }
}

# Delete old logs
function RemoveOldLogs($Original)
{
    try 
    {
        # Delete history from over a month ago
        $CleanedArray = @()
        for ($Row = 0; $Row -le $Original.Count - 1; $Row++)
        {
            $RelevantDates = ""
            $RowDates = $Original[$Row].dates.split(',') | % {if ($_[0] -eq " "){$_.Remove(0,1)} else {$_}} 

            # Only keep dates that are from the last 30 days
            foreach ($RowDate in $RowDates)
            {
                if (((get-date) - [convert]::ToDateTime($RowDate.Split(" ")[0])).days -le 30)
                {
                    $RelevantDates += $RowDate + ","
                }
            }

            # Save only relevant rows
            if ($RelevantDates -ne "")
            {
                $CleanedArray += $Row | select @{name="Volume"; Expression={$Original[$Row].Volume}}, 
                                               @{name="Count"; Expression={($RelevantDates.Split(",")).count - 1}}, 
                                               @{Name="Dates"; Expression={$RelevantDates.Remove($RelevantDates.Length - 1)}}
            }
        }

        return ($CleanedArray)
    }
    catch
    {
        return ($Original)
    }
}

import-module DataONTAP

# Set variables
$netapps = @("s-bch-na1","s-mif-na1")
$WantedVolSizePercent = 80
$MaxAggrPercent = 85

# Set credentials
$usr = "VolResizer"
$pass = "fujifilm1"
$encryptedPass = ConvertTo-SecureString -String $pass -AsPlainText -Force
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $usr,$encryptedPass

# Create log file
$path = "c:\VolResizer\logs\$(get-date -format "dd_MM_yyyy hh_mm").txt"
$log = New-Item -Path $path -Force -ItemType file 

# Delete old logs
Get-ChildItem  -path "c:\VolResizer\logs\*.txt" | where {$(get-date).Subtract($_.CreationTime).TotalDays -gt 7} | Remove-Item -Confirm:$false -Recurse

# Run on each netapp
foreach ($netapp in $netapps)
{
    # Connect to the current netapp
    Add-Content -Path $log -Value "Attemping to connect to $netapp"
    Connect-NcController $netapp -https -Credential $cred 

    # Check if log on was successful
    if ($global:CurrentNcController -eq $null)
    {
        Add-Content $log "Couldn't connect to $netapp"
    }
    else
    {
        # Get all SAN vols so we can check if there are old snapshots
        Add-Content $log "Attemping to get all vols"
        $vols = get-ncvol -Vserver *san* | ? {$_.VolumeIdAttributes.Type -ne "tmp" -and  $_.State -eq "online" -and 
                                              !$_.VolumeMirrorAttributes.IsDataProtectionMirror } | Sort-Object

        # Delete old snapshots
        Add-Content $log "Checking snapshots of all vols"
        foreach ($vol in $vols)
        {
            DeleteOldSnapshots($vol)
        }

        # Get vol increase history excel
        Add-Content $log "Importing vol increases csv"
        $IncreasedVols = Import-Csv "C:\VolResizer\$netapp.csv"
        $IncreasedVols = RemoveOldLogs($IncreasedVols)

        # Get volumes with used space of 85% or higher
        Add-Content $log "Getting volumes over 85%"
        $FullVols = get-ncvol -Vserver *san* -name !*mirror* | ? {($_.VolumeIdAttributes.Type -ne "tmp") -and ($_.Used -ge 85)} | Sort-Object TotalSize 

        # Check if you can increase the volume size
        foreach($FullVol in $FullVols)
        {
            Add-Content $log "Checking full volume $fullvol"
            $aggr = get-ncaggr $FullVol.Aggregate
            
            # Get the new vol size, how much is needed and whats the aggr capacity
            $NeededVolSizeGB = [math]::Ceiling((($FullVol.TotalSize / 1gb) * $FullVol.used / $WantedVolSizePercent) / 5) * 5
            $SpaceNeededGB = $NeededVolSizeGB - [math]::Round(($FullVol.TotalSize / 1gb),0)
            $AggrSizeGB = [math]::Round(($aggr.TotalSize / 1gb), 0)

            # Check conditions for increase
            if (($SpaceNeededGB -lt 300) -and ($aggr.used -lt 85) -and
                (($AggrSizeGB * $MaxAggrPercent / 100) -gt (($AggrSizeGB * $aggr.used / 100) + $SpaceNeeded)))
            {
                # Check if the vol wasnt increased over 10 times
                if  ($IncreasedVols | ? {($_.volume -eq $FullVol.name) -and ([convert]::ToInt32($_.count) -le 10)})
                { 
                    Add-Content $log "Increasing volume to $NeededVolSizeGB GB"
                    $IncreasedVols | ? {($_.volume -eq $FullVol.name) -and ($_.count.ToInt32($null) -le 15)} | % {$_.Count = [convert]::ToInt32($_.Count) + 1; $_.Dates = "$(get-date -Format "dd/MM/yyyy hh:mm:ss tt"),$($_.dates)"}
                    set-ncvolsize -Name $FullVol.name -VserverContext $FullVol.Vserver -NewSize ("$NeededVolSizeGB" + "GB")
                }
                # Check if the vol isnt in the table
                elseif (!($IncreasedVols | ? {($_.volume -eq $FullVol.name)}))
                {
                    Add-Content $log "Increasing volume to $NeededVolSizeGB GB"
                    $IncreasedVols += $FullVol | select @{name="Volume"; Expression={$FullVol.name}}, @{name="Count"; Expression={1}}, @{Name="Dates"; Expression={"$(get-date -Format "dd/MM/yyyy hh:mm:ss tt")"}}
                    set-ncvolsize -Name $FullVol.name -VserverContext $FullVol.Vserver -NewSize ("$NeededVolSizeGB"+"GB")
                }
            }
        }

        # Save Excel
        Add-Content $log "Exporting csv"
        $IncreasedVols | Export-Csv -LiteralPath "C:\VolResizer\$netapp.csv" -Force:$true 
        $global:CurrentNcController = $null
    }
}