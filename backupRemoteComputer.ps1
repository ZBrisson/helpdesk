# Back Up User Profile Information
#By: Zach Brisson
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
# Remote App Data File Location
$roamingAppData = "\\$sourceComputer\c$\Users\$username\AppData\Roaming"

function copyAll() {
    copyMembership
    copyOutlook
    copyPrinter
    copyStartMenu
    copyTaskBar
}

function copyMembership() {

    Get-ADComputer $sourceComputer -Properties MemberOf | Export-CSV -path "$backupDirectory\ADMembership_export.csv"
}

function copyOutlook() {


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

    if (Test-Path $filePath) {
        Write-Host "Copying Taskbar favorites from $sourceComputer to $backupDirectory\Taskbar"
        Copy-Item $filePath -Destination "$backupDirectory\TaskBar" -Recurse
        Write-Host "Successfully copied Taskbar favorites information." -ForegroundColor Green
    }
    else {
        Write-Host "No Folder Found. Are there any programs pinned to the taskbar?" -ForegroundColor Yellow
    }
}

