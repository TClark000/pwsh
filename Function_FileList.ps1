#Function generates a list of filesnames within a path, returns array $FileNames
Function FileList {

param(
[Parameter(position=1,mandatory=$true)]
[ValidateScript({Test-Path $_ -PathType 'Container'})]
[string]$Path = ""
)

process{

[array]$script:FileNames = @()
[array]$FileName = @()

Get-ChildItem $Path | Where-Object {$_.PSIsContainer -eq $False} | ForEach-Object {
    $FileName = New-Object PSObject
    $FileName | Add-Member NoteProperty Name $_.Name
    $script:FileNames += $FileName
    }

    }

    }
