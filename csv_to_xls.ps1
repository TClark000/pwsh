#Appends multiple csv files from $PathFiles that are named with a matching string from $SearchNames into a single corresponding named csv
#Appended csv are saved into $PathResult (overrides corresponding files)
#Appended csv then copied into a single xls, 1 csv per worksheet, resulting ss named $PathResultFile
#Function required 'Function_FileList.ps1'

$error.clear()
$ErrorActionPreference = "Stop"
Clear-Host

#Query path of csv files
[string]$PathFiles = "C:\temp"
$Msg = "The default csv path is " + $PathFiles + ".  Enter a new path or press return for the default shown"
$ResponsePath = Read-Host -prompt $Msg
If($ResponsePath){
    if(!(Test-Path $ResponsePath -pathType container)){
    Write-Output "File path is invalid."
    break }
    $PathFiles = $ResponsePath}

#query csv name search criteria
[array]$SearchNames = @("Installed","Admins","Users")
[array]$ResponseSearch =@()
$Msg = "Enter a new csv search criteria, separate by commas or press return for the default shown -`n" + ($SearchNames -join ", ")
$ResponseSearch = Read-Host $Msg
If($ResponseSearch){
$SearchNames = invoke-expression "write-output $ResponseSearch"
[array]::sort($SearchNames)
}

[string]$PathResult = $PathFiles + "\Result"
[string]$CurrentPath = @()
[string]$SearchName = @()
[array]$FilesCollate = @()
[array]$CSVCollate = @()
[array]$FoundNames = @()
[array]$FoundName = @()
[Int32]$CSVCollateCount = 0
[string]$PathResultFile = $PathResult + "\Report.xlsx"
[Int32]$LoopCSV = 0
[string]$WorksheetName =""
[string]$Msg = ""

#--------------------------
#Check search criteria
$WshShell = New-Object -ComObject wscript.shell
$Message ="Append csv files from " + $PathFiles + ", save results to " + $PathResult + " based on the search of: `n" + ($SearchNames -join ", ")
$PopUp = $WshShell.popup("$Message",0,"PowerShell Script Confirmation",1)
if ($PopUP -eq 2){
break}

#Create result directory if required; Out-Null verbose quiet
If(!(Test-Path $PathResult -pathType container)){New-Item -Path $PathFiles -Name "Result" -ItemType "directory" | Out-Null }

#List the file names that are in $PathFiles directory into $FileNames - from 'Function_FileList.ps1'
FileList -path $PathFiles

#Compare $FileNames with each item in $SearchNames
#For each $SearchName match append/import csv and then export csv to $PathResult
foreach ($SearchName in $SearchNames) {

    [array]$FilesCollate = @()
    [array]$CSVCollate = @()

    $FilesCollate = $FileNames | Sort-Object Name | Where-Object {$_ -match $SearchName -and $_ -match "csv"}

    if ($FilesCollate) {

        $FilesCollate | ForEach-Object{

            $CurrentPath = $Pathfiles + "\" + $_.Name
            $CSVCollate += import-csv -Path $CurrentPath
            $CSVCollateCount += 1
            }

        $CurrentPath = $PathResult + "\" + $SearchName + ".csv"
        $CSVCollate | export-csv -Path $CurrentPath -NoTypeInformation

        $FoundName = New-Object PSObject
        $FoundName | Add-Member -type NoteProperty Name $SearchName
        $FoundNames += $FoundName

        }
}

#If no csv files break script
If (!$CSVCollateCount){
    $Msg = "No csv files in the path " + $PathFiles + " that match the strings:`n" + ($SearchNames -join ", ")
    Write-Output $Msg
    break}
$Msg = "Unappended csv files total: " + $CSVCollateCount
Write-Output $Msg
$Msg = "Appended csv files total: " + $FoundNames.count
Write-Output $Msg

#Create Excel Object and add workbook/worksheet (newest created worksheet is always 1)
$excel = new-object -ComObject excel.application
$excel.DisplayAlerts = $True
$excel.Visible = $False
$workbook = $excel.workbooks.Add()
$worksheet = $workbook.worksheets.Item(1)

#collate\loop each csv and save into a single excel workbook
$FoundNames| Sort-Object name -Descending | foreach-object {$_} {

        $LoopCSV += 1
        $CurrentPath = $PathResult + "\" + $_.name + ".csv"

        #Open and copy contents of the CSV file
        $tempcsv = $excel.Workbooks.Open($CurrentPath)
        $tempsheet = $tempcsv.Worksheets.Item(1)
        $tempSheet.UsedRange.Copy() | Out-Null

        #Paste contents of CSV into worksheet
        $worksheet.Paste()

        #Name worksheet
        #$WorksheetName = ($_.name -replace ".{4}$" )
        $WorksheetName = ($_.name)

        $worksheet.Name = $WorksheetName

        #Check if another worksheet required
        if ($FoundNames.Count -gt $LoopCSV){
            $worksheet = $workbook.worksheets.add()
            }
}

$workbook.saveas($PathResultFile)
$Excel.Workbooks.Close()
$excel.quit()
$Msg = "Resulting ss saved as: " + $PathResultFile
Write-Output $Msg
$error
