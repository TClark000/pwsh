#Function generates a list of filesnames within a path, returns array $FileNames
Function FileList {

param(
[Parameter(position=1,mandatory=$true,HelpMessage="Default path is set to: C:\Temp")]
[ValidateScript({Test-Path $_ -PathType 'Container'})]
[string]$Path = 'C:\Temp'
)

process{

[array]$script:FileNames = @()
[array]$FileName = @()

Get-ChildItem $Path | ? {$_.PSIsContainer -eq $False} | % {
    $FileName = New-Object PSObject
    $FileName | Add-Member NoteProperty Name $_.Name
    $script:FileNames += $FileName
    }

    }
    
    }