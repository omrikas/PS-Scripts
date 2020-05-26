param (
    [Parameter(Mandatory=$true)][string]$FI
)

Import-Module *Cisco.UCSManager*

# set credentials
$usr = read-host "Enter Username: "
$pass = read-host "Enter Password: " -AsSecureString
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $usr,$pass 

# Connect to the FI 
Connect-Ucs $FI -Credential $cred

# Check if successfully connected
if ((get-ucspssession) -ne $null)
{
    # Get all the faults
    $UCSServers = Get-UcsServer | Sort-Object ServerId | Select-Object AssignedToDn, ChassisId, SlotId, AvailableMemory

    # Create a new Excel
    $Excel = New-Object -ComObject excel.application
    $WorkBook = $Excel.Workbooks.add()
    $WorkSheet = $WorkBook.Sheets.Item(1)
    $WorkSheet.DisplayRightToLeft = $true
    $Excel.Visible = $true
    $Excel.WindowState = "xlMaximized"
    
    # Starting location for chassis map
    $Row = 2
    $Col = 2
    
    # Starting location for Sum map
    $SumRow = 10
    $SumCol = 2

    # Esxi clusters (replace a-g with fitting names, if you have more/lss clusters add/remove rows and later on in the switch case accordingly)
    $WorkSheet.Cells.Item(11, 1).value() = "a"
    $WorkSheet.Cells.Item(12, 1).value() = "b"
    $WorkSheet.Cells.Item(13, 1).value() = "c"
    $WorkSheet.Cells.Item(14, 1).value() = "d"
    $WorkSheet.Cells.Item(15, 1).value() = "e"
    $WorkSheet.Cells.Item(16, 1).value() = "f"
    $WorkSheet.Cells.Item(17, 1).value() = "g"
    $WorkSheet.Cells.item(10, 1).interior.colorindex = 1; $WorkSheet.Cells.item(10, 1).BorderAround(1,3,1) *> $null
    $WorkSheet.range($WorkSheet.cells.item(11, 1), $WorkSheet.cells.item(15, 1)).BorderAround(1,3,1) *> $null
    $WorkSheet.range($WorkSheet.cells.item(10, 2), $WorkSheet.cells.item(10, 1 + ($UCSServers | Group-Object ChassisId).count)).BorderAround(1,3,1) *> $null
    $WorkSheet.range($WorkSheet.cells.item(11, 2), $WorkSheet.cells.item(15, 1 + ($UCSServers | Group-Object ChassisId).count)).value = 0

    # Map each chassis
    foreach ($CurrChassis in ($UCSServers | Group-Object ChassisId))
    {
        # Chassis Title
        $WorkSheet.Cells.Item($Row, $Col).value() = "Chassis " + $CurrChassis.name
        $WorkSheet.Cells.item($Row, $Col).font.bold = $true  
        $WorkSheet.range($WorkSheet.cells.item($Row, $Col), $WorkSheet.cells.item($Row, $Col + 1)).MergeCells = $true
        $WorkSheet.range($WorkSheet.cells.item($Row, $Col), $WorkSheet.cells.item($Row, $Col + 1)).BorderAround(1,3,1) *> $null
        $WorkSheet.Cells.item($Row, $Col).horizontalalignment = -4108

        $WorkSheet.Cells.Item($SumRow, $SumCol).value() = "Chassis " + $CurrChassis.name

        # Chassis cells
        for ($i = 1; $i -lt 9; $i++)
        {
            # Set current cell location
            $CellRow = $Row + [math]::Ceiling($i / 2)
            $CellCol = $Col + 1 - $i % 2 

            # Create Current Cell
            $WorkSheet.Cells.item($CellRow, $CellCol).BorderAround(1,3,1) *> $null
            $WorkSheet.Cells.item($CellRow, $CellCol).horizontalalignment = -4108

            # Check if current cell has blade
            $CellData = $CurrChassis.Group | ? {$_.slotid -eq $i}

            # If current cell has a blade, print its profile
            if ($CellData)
            {
                $ServerProfile = $CellData.AssignedToDn.split('/')[2].remove(0,3)
                $WorkSheet.Cells.item($CellRow, $CellCol).value() = $ServerProfile

                # Check if the current server is the size of 2 slots
                if ($CellData.AvailableMemory -and ($CellData.AvailableMemory -gt 900000))
                {
                    $WorkSheet.range($WorkSheet.cells.item($CellRow, $CellCol), $WorkSheet.cells.item($CellRow, $CellCol + 1)).MergeCells = $true
                    $WorkSheet.range($WorkSheet.cells.item($CellRow, $CellCol), $WorkSheet.cells.item($CellRow, $CellCol + 1)).BorderAround(1,3,1) *> $null
                }
                
                # Check if current server is an esxi
                if ($ServerProfile.Contains("esxi"))
                {
                    # Map esx type by its name (change a-g to fitting categories and the switch conditions to what fits the enviorment)
                    switch ($ServerProfile.Split("-")[2].split("esx")[0])
                    {
                        # a
                        "a"   {$WorkSheet.Cells.Item($SumRow + 1, $SumCol).value() = $WorkSheet.Cells.Item($SumRow + 1, $SumCol).value() + 1}
                        
                        # b
                        "b"  {$WorkSheet.Cells.Item($SumRow + 2, $SumCol).value() = $WorkSheet.Cells.Item($SumRow + 2, $SumCol).value() + 1}

                        # c
                        "c"  {$WorkSheet.Cells.Item($SumRow + 3, $SumCol).value() = $WorkSheet.Cells.Item($SumRow + 3, $SumCol).value() + 1}

                        # d
                        "d" {$WorkSheet.Cells.Item($SumRow + 4, $SumCol).value() = $WorkSheet.Cells.Item($SumRow + 4, $SumCol).value() + 1}

                        # e
                        "e"  {$WorkSheet.Cells.Item($SumRow + 5, $SumCol).value() = $WorkSheet.Cells.Item($SumRow + 5, $SumCol).value() + 1}

                        # f
                        "f"  {$WorkSheet.Cells.Item($SumRow + 6, $SumCol).value() = $WorkSheet.Cells.Item($SumRow + 6, $SumCol).value() + 1}

                        # g
                        "g"  {$WorkSheet.Cells.Item($SumRow + 7, $SumCol).value() = $WorkSheet.Cells.Item($SumRow + 7, $SumCol).value() + 1}
                    }
                }
            }

            $range = $WorkSheet.UsedRange; $range.EntireColumn.AutoFit() *> $null;
        }

        $SumCol += 1
        $col = $col + 3
    }

    Disconnect-Ucs
}
else
{
    "Connection Failed"
}
 