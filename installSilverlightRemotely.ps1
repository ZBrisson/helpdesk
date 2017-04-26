# Silverlight updater
# By: Zach Brisson
# 04/26/17
# Script will run through a list of computers and update them to whatever version
# is at the 32 and 64 path. 
# Can do single computer by uncommenting $list_of_computers

## get the list of computers from the text file
$list_of_computers = get-content "\\remoteserverhere\computerlist.txt"
# $list_of_computers = (Get-ADComputer -SearchBase 'OU=Workstations,OU=Palm Beach Gardens,OU=Branch Court Services,OU=Operations,OU=CLERK,DC=Clerk,DC=local' -Filter '*' -Properties Name).Name
# $list_of_computers = $list_of_computers.Trim()
# $list_of_computers = Read-Host "Type the name of a PC"

#display total count
Write-host "Gathered machines" $list_of_computers.count
# Silverlight .exe 32bit file path
$silverlight32Path = "\\remoteserverhere\\Silverlight.exe"
# Silverlight .exe 64bit file path
$silverlight64Path = "\\\\remoteserverhere\\Silverlight_x64.exe"
#The .exe resting place on the remote computer
$installLocation = "\\$computer\c$\Software\"


#Install Silverlight
 foreach ($computer in $list_of_computers) 
    {

        # ping the computer and check the operating system if a pong is heard, otherwise (else) notify the pc is not responding
        if (Test-Connection $computer -count 1 -Quiet)
        {
                ## check the Windows OS architecture
            if ((Get-WmiObject -ComputerName $computer win32_operatingsystem | select osarchitecture).osarchitecture -eq "64-bit")
            {

                Write-Host "$computer is a 64-bit OS" -ForegroundColor Green
                Copy-Item $silverlight64Path -Destination $installLocation -Recurse
                Write-Host "Copied silverlight successfully at $installLocation" -ForegroundColor Green
                $InstallString = "$installLocation\Silverlight_x64.exe /q"

        
            }else{
        
                Write-Host "$computer is a 32-bit OS" -ForegroundColor Green
                Copy-Item $silverlight32Path -Destination $installLocation -Recurse
                Write-Host "Copied silverlight successfully $installLocation" -ForegroundColor Green
                $InstallString = "$installLocation\Silverlight.exe /q"

            }
            # Install silverlight
            ([WMICLASS]"\\$computer\ROOT\CIMV2:Win32_Process").Create($InstallString) 
            Write-Host "Installed silverlight successfully" -ForegroundColor Green

        }
        else{

            # notify on screen that host is offline
            Write-Host "$computer is offline, skipping..." -ForegroundColor Yellow

        }

     }

Write-Host ""
Write-Host "Install script completed. Users will need to restart IE for changes to update." -ForegroundColor Green
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "Going to sleep for 3 minutes to let install finish and then delete the files in \\ComputerName\C$\Software" -ForegroundColor Yellow

#Remove downloaded files from remote computers
Start-Sleep 180
 foreach ($computer in $list_of_computers)  {
            Remove-Item $installLocation -recurse
            Write-Host "Removed .exe from $installLocation" -ForegroundColor Green
}
