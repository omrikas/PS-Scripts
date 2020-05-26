# sweep based on a subnet of 255.255.255.0
$segment = Read-Host "please enter segment (aaa.bbb.ccc)"
1..255 | % {"$segment.$($_): $(Test-Connection -count 1 -ComputerName "$se1gment.$($_)" -quiet)"}