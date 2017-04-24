#Copy user's taskbar and start menu.
##### Zach Brisson ######  04/23/17
# Links must be the same. 
#I.E. Copying the shortcuts for a 32 bit computer will not work when transferring to 64 bit computer without modification(adding the x86 file path in the shortcuts). 





Param(
    [Parameter(Mandatory=$True)]
    [string]$username
)


function copyTaskBar($username) {

    $filePath = "C:\Users\$username\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

    if (Test-Path $filePath) {
    Copy-Item $filePath -Destination "C:\$username\TaskBar" -Recurse
    }
    else {
        Write-Host "No Folder Found. Are there any programs pinned to the taskbar?"
    }
}

function copyStartMenu($username) {

    $filePath = "C:\Users\$username\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu"

    if (Test-Path $filePath) {
    Copy-Item $filePath -Destination "C:\$username\StartMenu" -Recurse
    }
    else {
        Write-Host "No Folder Found. Are there any programs pinned to the Start Menu? Windows 8 and up stores the start menu diffierently."
    }

}
