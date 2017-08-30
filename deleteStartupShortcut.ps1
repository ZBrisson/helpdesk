#-----------------------------------------------------------------------------
# Script: deleteStartupShortcut
# Author: Zach Brisson
# Date: 08/30/17
# Searches each user folder for a specified computer and deletes a shortcut entry.
# Also deletes the all users shortcut for all future users. 
# About half of this code has been sourced from a previous script made by Corey Wade


## get the list of computers from the text file
$list_of_computers = get-content C:\scripts\deleteStartupShortcut\pcs.txt
# $list_of_computers = (Get-ADComputer -SearchBase 'OU=<FILLHERE>,DC=<FAKEOUTEXT>,DC=local' -Filter '*' -Properties Name).Name
# $list_of_computers = $list_of_computers.Trim()
# $list_of_computers = "<COMPUTERNAME>"


#setup errorcode variable
$errorcode = 0

#print error code
$PrintErrorcode = {
Write-Host "`n"
Write-Host "Exiting with code $errorcode"
.$endpause
}

#setup endpause
$endpause = {
Write-Host "`n"
write-host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-host
}


foreach ($pcname in $list_of_computers)
{
    if (Test-Connection $pcname -count 1 -quiet) {

        #setup the userfolders variable to pull the list of users from c:\users of the target machine
        $SetupVariables = {
            if ($errorcode -ne 0) {
                break;
            }
            write-host "`n"
            $userfolders = Get-ChildItem -Path \\$pcname\c$\Users
        }

        .$SetupVariables
    
        Write-Host "$pcname is online" -Foreground "green"
        Write-Host "=- Removing from $pcname"
        #foreach loop which performs the actual deletions
         $EraseFolders = {
        if ($errorcode -ne 0) {
            break;
        }
            foreach ($userfolder in $userfolders)
                {
                Write-Host "=- Removing from $userfolder"
                Remove-Item -Recurse -Force -erroraction 'silentlycontinue' "\\$pcname\c$\Users\$userfolder\AppData\Roaming\Microsoft\Windows\Start Menu\<SHORTCUT LOCATION HERE>*"
                Remove-Item -Recurse -Force -erroraction 'silentlycontinue' "\\$pcname\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\<SHORTCUT FOLDER>\<SHORTCUT LOCATION HERE>*"}
                }
        .$EraseFolders
        .$PrintErrorcode

    }else{
    
        write-host "$pcname is not responding. Stopping script." -foreground "red"
        write-host "`n"
        Write-Host "Please ensure that this PC is powered on and connected to the CLERK network." 
    
    }

}
