# Back Up User Profile Information
# By: Zach Brisson
# 04/24/2017
# Copy user information that is not networked with the cloud drives.



Param(
    [Parameter(Mandatory=$True)]
    [string]$sourceComputer,

    [Parameter(Mandatory=$True)]
    [string]$username

)

# Backup Directory
$backupDirectory = "C:\$username"
# Local Computer
$localComputer = $env:COMPUTERNAME
#PS Remote
$psremote = New-PSSession -ComputerName $sourceComputer
# Remote App Data File Location
$roamingAppData = "\\$sourceComputer\c$\Users\$username\AppData\Roaming"

#get the sid number for a username
$email = "$username@mypalmbeachclerk.com" 
$sid = (New-Object System.Security.Principal.NTAccount($email)).Translate([System.Security.Principal.SecurityIdentifier]).value

function copyAll() {
    copyMembership
    copyOffice
    copyPrinter
    copyStartMenu
    copyTaskBar
}

function copyMembership() {

    Get-ADComputer $sourceComputer -Properties MemberOf | Export-CSV -path "$backupDirectory\ADMembership_export.csv"
}

function copyOffice() {


    Write-Host "Copying Office Information from $sourceComputer to $backupDirectory"
    # Outlook Roaming Settings. 
    $filePath = "$roamingAppData\Microsoft\Outlook"
    if (Test-Path $filePath) {
        Copy-Item $filePath -Destination "$backupDirectory\Outlook\Roaming" -Recurse
    }
    # Office Stationary
    $filePath = "$roamingAppData\Microsoft\Stationary"
    if (Test-Path $filePath) {
        Copy-Item $filePath -Destination "$backupDirectory\Stationary" -Recurse
    }
    # Office Signatures
    $filePath = "$roamingAppData\Microsoft\Signatures"
        if (Test-Path $filePath) {
        Copy-Item $filePath -Destination "$backupDirectory\Signatures" -Recurse
    }
    # Office Templates
    $filePath = "$roamingAppData\Microsoft\Templates"
        if (Test-Path $filePath) {
        Copy-Item $filePath -Destination "$backupDirectory\Templates" -Recurse
    }
    # Office Dictionary 
    $filePath = "$roamingAppData\Microsoft\Proof"
        if (Test-Path $filePath) {
        Copy-Item $filePath -Destination "$backupDirectory\Proof" -Recurse
    }
    Write-Host "Successfully copied Office information." -ForegroundColor Green
}

function copyPrinter() {
    Write-Host "Copying Printer Information from $sourceComputer to $backupDirectory\printer_export.csv"
    Write-Host "***Tip:Local USB Printers are named. Networked Printers have \\ in front of their name***"
    # Grab Printer Setup and export to .csv
    Get-WMIObject -class Win32_Printer -computer $sourceComputer | Select Name,Location,SysteName | Export-CSV -path "$backupDirectory\printer_export.csv"
    Write-Host "Successfully copied Printer information." -ForegroundColor Green
}

function copyStartMenu($username) {

    $filePath = "$roamingAppData\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu"

    if (Test-Path $filePath) {
        Write-Host "Copying Start Menu favorites from $sourceComputer to $backupDirectory\StartMenu"
        Copy-Item $filePath -Destination "$backupDirectory\StartMenu" -Recurse
        Write-Host "Successfully copied Start Menu favorites information." -ForegroundColor Green
    }
    else {
        Write-Host "No Folder Found. Are there any programs pinned to the Start Menu? Windows 8 and up stores the start menu diffierently." -ForegroundColor Yellow
    }

}

function copyTaskbar() {
    $filePath = "$roamingAppData\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    $regFilePath = "HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    $regCmd = "export"
    $regExport = { Param ($RegCmd,$regFilePath,$backupDirectory) Invoke-expression "C:\windows\system32\reg.exe $RegCmd $RegFilePath c:\temp\taskbar.reg -y"}

    if (Test-Path $filePath) {
        Write-Host "Copying Taskbar favorites from $sourceComputer to $backupDirectory\Taskbar"
        Copy-Item $filePath -Destination "$backupDirectory\TaskBar" -Recurse
        Invoke-Command -Session $psremote -ArgumentList $RegCmd,$RegFilePath -scriptblock $regExport
        Copy-Item "\\$sourceComputer\c$\temp\taskbar.reg" -Destination "$backupDirectory\Taskbar\taskbar.reg" -Recurse
        Remove-Item -Recurse "\\$sourceComputer\c$\temp\taskbar.reg"
        Write-Host "Successfully copied Taskbar favorites information." -ForegroundColor Green
    }
    else {
        Write-Host "No Folder Found. Are there any programs pinned to the taskbar?" -ForegroundColor Yellow
    }
}


function Show-Menu
{
     param (
           [string]$Title = 'Backup Remote Computer Settings'
     )
     Write-Host "================ $Title ================"
    
     Write-Host "copy All: Press '1' for this option."
     Write-Host "copy Active Directory Membership info: Press '2' for this option."
     Write-Host "copy Office: Press '3' for this option."
     Write-Host "copy Printer Info: Press '4' for this option."
     Write-Host "copy Start Menu: Press '5' for this option."
     Write-Host "copy Taskbar Menu: Press '6' for this option."
     Write-Host "Q: Press 'Q' to quit."
}

do
{
     Show-Menu
     $input = Read-Host "Please make a selection"
     switch ($input)
     {
           '1' {
                cls
                'You chose option #1'
                copyAll
           } '2' {
                cls
                'You chose option #2'
                copyMembership
           } '3' {
                cls
                'You chose option #3'
                copyOffice
           }'4' {
                cls
                'You chose option #4'
                copyPrinter
           } '5' {
                cls
                'You chose option #5'
                copyStartMenu
           } '6' {
                cls
                'You chose option #6'
                copyTaskbar

           } 'q' {
                return
           }
     }
     pause
}
until ($input -eq 'q')
