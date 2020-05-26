# Recursively get all params of the object
function GetParams([object]$Object, [string]$currPath)
{
    # Check if object exists and the current depth isnt too big
    if ($Object -and (([regex]::Matches($currPath, ".")).count -lt $MaxDepth))
    {
        # Run on all params of current object
        foreach ($member in ($Object | get-member | ? {$_.MemberType -eq 'Property'}))
        {
            # get current propery
            $CurrProperty = $Object.$($member.Name)

            # Check if curr property isnt null
            if ($CurrProperty)
            {
                # Check if the table already contains this value (to stop infinite loop)
                if  (($Table.Values | ? {$_ -eq $CurrProperty}).Count -le 1)
                {
                    # Try to add the value to the table
                    try {$Table.add(($currPath + "." + $member.Name), $CurrProperty)} catch{}
                    GetParams $CurrProperty ($currPath + "." + $member.Name);
                }
            }
        }
    }
}

# Params
$MaxDepth = 15
$Table = @{}

#Example
$t = Get-Host
GetParams($t, "")
