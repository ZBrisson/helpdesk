# Silverlight updater
# By: Zach Brisson
# 04/26/17
# You need to have the most up to date version in the silverlight folder on the Z drive.
# 

## get the list of computers from the text file
#$list_of_computers = get-content "C:\computerlist.txt"
# $list_of_computers = (Get-ADComputer -SearchBase 'OU=Workstations,OU=Palm Beach Gardens,OU=Branch Court Services,OU=Operations,OU=CLERK,DC=Clerk,DC=local' -Filter '*' -Properties Name).Name
# $list_of_computers = $list_of_computers.Trim()
$computer = "COC-V-ZJB1000"
# Read-Host "Type the name of a PC"

$fileVersion = (Get-ItemProperty -Path "\\wcp01zfs-03.clerk.local\IT\Software\Silverlight\Silverlight_x64.exe").VersionInfo.FileVersion
$installVersion = if (Test-Path "\\$computer\c$\Program Files\Microsoft Silverlight\") {
    (Get-ItemProperty -Path "\\$computer\c$\Program Files\Microsoft Silverlight\").VersionInfo.FileVersion
}else {
    (Get-ItemProperty -Path "\\$computer\c$\Program Files (x86)\Microsoft Silverlight\").VersionInfo.FileVersion
}
$installFolder = "Microsoft Silverlight"
$remoteSession = New-PSSession -computername $computer

# ping the computer and check the operating system if a pong is heard, otherwise (else) notify the pc is not responding
if ((Test-Connection $computer -count 1 -Quiet) -and ($installVersion -le $fileVersion))
    {
        ## check the Windows OS architecture
        if ((Get-WmiObject -ComputerName $computer win32_operatingsystem | select osarchitecture).osarchitecture -eq "64-bit")
            {

                $installLocation = "\\$computer\c$\Program Files (x86)\$installFolder"  
                $silverlightDownload = "\\wcp01zfs-03.clerk.local\IT\Software\Silverlight\Silverlight_x64.exe"
                Write-Host "$computer is a 64-bit OS" -ForegroundColor Green 
                Copy-Item $silverlightDownload -Destination $installLocation -Force
                Invoke-Command -session $remoteSession -scriptblock {'C:\Program Files (x86)\Microsoft Silverlight\Silverlight_x64.exe /q'}

            }else{

                $installLocation = "\\$computer\c$\Program Files\$installFolder"
                $silverlightDownload = "\\wcp01zfs-03.clerk.local\IT\Software\Silverlight\Silverlight.exe"
                Write-Host "$computer is a 32-bit OS" -ForegroundColor Green
                Copy-Item $silverlightDownload -Destination $installLocation -Force
                Invoke-Command -session $remoteSession -scriptblock {'C:\Program Files\Microsoft Silverlight\Silverlight.exe /q'}

            }
    }
        
else{

        # notify on screen that host is offline
        Write-Host "$computer is offline or silverlight is up to date, skipping..." -ForegroundColor Yellow

}


Write-Host ""
Write-Host "Install script completed. Users will need to restart IE for changes to update." -ForegroundColor Green
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "Going to sleep for 3 minutes to let install finish and then delete the files" -ForegroundColor Yellow

#Remove downloaded files from remote computers
Start-Sleep 180
 foreach ($computer in $list_of_computers)  {
            Remove-Item $installLocation
            Write-Host "Removed .exe from $installLocation" -ForegroundColor Green
}
