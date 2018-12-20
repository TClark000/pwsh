#Retrieves NIC details - Name, MAC and Speed.
Function GetMac {
Param (
[Parameter(Mandatory=$true)]
[string[]] $computername
)
Process{
[string]$pattern1 = "Ethernet"
[string]$pattern2 = "Virtual"
[string]$pattern3 = "PCI"
[string]$Msg = ""
if (test-connection -computername $computername -quiet -count 1) {
gwmi win32_NetworkAdapter -ComputerName $computername | Select @{l='WorkStation';e={$computername}},Name, @{l='MAC Address';e={$_.MACAddress}}, @{l='Max Speed (Mb)'; e={$_.Speed / 1000000 -as [int]}}
}
Else {
$Msg = $computername + "is offline"
write-output $Msg
}
}
}
