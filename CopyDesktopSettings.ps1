# Script to copy the desktop settings like text scalling.
# saves as a .reg file.

#connect to the main module which uses loop and lists.
Param($sourceComputer,$remoteComputer)

function copyDesktopSettings(){
#File Location
$file1 = \yourfilesharefile1.txt

#Check for Existence of File
$file1Exists = Test-Path $file1

    If ($file1Exists)
    {
        $confirmation = Read-Host "File already exists. Do you want to override? (y or n)"
            If ($confirmation -eq 'y')
            {
                Reg export 'HKCU\Control Panel\Desktop' C:\DesktopSettings.reg /y
            }
    }
    else {
        Reg export 'HKCU\Control Panel\Desktop' C:\DesktopSettings.reg
        }
}

# How to paste the values to the target computer. 
# Reg Import exportedkey.reg
