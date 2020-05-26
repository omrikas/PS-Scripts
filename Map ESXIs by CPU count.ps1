Import-Module VMware.VimAutomation.Core

# Create the excel with the filled cells
function CreateExcel()
{
    # Write cells text
    $WorkSheet.Cells.Item(1, 1).value() = 'VCenter'    
    $WorkSheet.Cells.Item(1, 2).value() = 'Cluster' 
    $WorkSheet.Cells.Item(1, 3).value() = 'Name' 
    $WorkSheet.Cells.Item(1, 4).value() = 'CPU' 

    # Set the cells as bold and text in center
    for ($i = 1; $i -le 4; $i = $i + 1)
    {
        if ($i -lt 6)
        { $MergeCells = $WorkSheet.range($WorkSheet.cells.item(1,$i), $WorkSheet.cells.item(2,$i))
          $MergeCells.MergeCells = $true
          $mergeCells.BorderAround(1,3,1) *> $null }
        $WorkSheet.Cells.item(1, $i).font.bold = $true  
        $WorkSheet.Cells.item(1, $i).horizontalalignment = -4108
        $WorkSheet.Cells.item(1, $i).verticalalignment = -4108
    }

    # Auto fit
    $range = $WorkSheet.UsedRange; $range.borderAround(1,3,1); $range.EntireColumn.AutoFit();
}

# Create a new Excel
$Excel = New-Object -ComObject excel.application
$WorkBook = $Excel.Workbooks.add()
$WorkSheet = $WorkBook.Sheets.Item(1)
$WorkSheet.DisplayRightToLeft = $false
$Excel.Visible = $true
CreateExcel *> $null
$CurrentRow = 3

# Get all vCenter to map
$vcenters = (read-host "Enter all vcenters to map (with a ,) between each name").split(',').Replace(" ", "")

# Map each vcenter
foreach ($vcenter in $vcenters)
{
    # Connect to vCenter- works best if the user running this script has permissions on the vCenter
    Connect-VIServer $vcenter
    $VMhosts = Get-VMHost | Sort-Object
    $StartingRow = $CurrentRow
   
    # Map all hosts
    foreach ($VMhost in $VMhosts)
    {
        $WorkSheet.Cells.Item($CurrentRow, 1).value() = $vcenter
        $WorkSheet.Cells.Item($CurrentRow, 2).value() =  (Get-Cluster -VMHost $VMhost).name
        $WorkSheet.Cells.Item($CurrentRow, 3).value() = $VMhost.name
        $WorkSheet.Cells.Item($CurrentRow, 4).value() = $vmhost.NumCpu
        $CurrentRow = $CurrentRow + 1
    }

    # Auto fit
    $range = $WorkSheet.UsedRange; $range.borderAround(1,3,1); $range.EntireColumn.AutoFit();

    # Merge all the cells of the vcenter
    $MergeCells = $WorkSheet.range($WorkSheet.cells.item($StartingRow,1), $WorkSheet.cells.item($CurrentRow - 1,1))
    $MergeCells.MergeCells = $true; $mergeCells.BorderAround(1,3,1) *> $null ; 
    $WorkSheet.Cells.item($CurrentRow - 1,1).horizontalalignment = -4108; $WorkSheet.Cells.item($CurrentRow - 1,1).verticalalignment = -4108

    Disconnect-VIServer -Confirm:$false *>$null
}