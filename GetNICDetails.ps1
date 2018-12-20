#Function required 'Function_GetMac.ps1'
#Collate workstation MAC addresses from LAN\ethernet NICs.  (Collect info for SCCM builds.)

[string]$FilePaths = "c:\temp\test\"
[string]$csvFile = $FilePaths + "result.csv"
[string]$WrkStFile = $FilePaths + "active.txt"
[array]$WorkStations =@()
[array]$WkSDetails =@()
[array]$NICDetails = @()
[array]$NameOption = @("VPN","Wireless")
[bool]$NicBool = 0
$WorkStations = get-content -path $WrkStFile
ForEach ($_ in $WorkStations){
    $WkSDetails = GetMAC $_
    $NicBool = 0
    ForEach ($_ in $WksDetails){
    $NicDetail = New-Object PSObject
    If ( ($_.'Max Speed (Mb)' -gt 60) -and ($_.Name -notmatch $NameOption[0]) -and ($_.Name -notmatch $NameOption[1])){
    $NicDetail | Add-Member NoteProperty WorkStation $_.WorkStation
    $NicDetail | Add-Member NoteProperty Name $_.Name
    $NicDetail | Add-Member NoteProperty 'MAC Address' $_.'MAC Address'
    $NicDetail | Add-Member NoteProperty 'Max Speed (Mb)' $_.'Max Speed (Mb)'
    $NicDetails += $NicDetail
    $NicBool = 1
    }
    }
If(-not $NicBool) {$NicDetails += $WksDetails}
}
$NICDetails | format-table
$NICDetails | Export-csv -Path ($csvFile) -NoTypeInformation
