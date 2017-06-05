# Axia Install
# By: Zach Brisson
# 05/30/17
# You will need to manually install the dependcies


## get the list of computers from the text file
#$list_of_computers = get-content "C:\computerlist.txt"
# $list_of_computers = (Get-ADComputer -SearchBase 'OU=text,OU=dummy,DC=blah,DC=local' -Filter '*' -Properties Name).Name
# $list_of_computers = $list_of_computers.Trim()
$list_of_computers = Read-Host "Type the name of a PC"

#display total count
Write-host "Gathered machines" $list_of_computers.count
# Axia File Path
$axiaPath = "\\network\file\place\holder\text"
$axiaTestPath = "\\network\file\place\holder\text"
$crystalPath = "\\network\file\place\holder\text"
$axiaIcon = "\\network\file\place\holder\text"
$axiaTestIcon = "\\network\file\place\holder\text"


$check_arch =
{

    ## check the Windows OS architecture
    if ((Get-WmiObject -ComputerName $computer win32_operatingsystem | select osarchitecture).osarchitecture -eq "64-bit")
    {

       
        Write-Host "64-bit OS" -ForegroundColor Green

        $program_files = "Program Files (x86)"

        $os_arch = "64"
        
    }else{
        
        Write-Host "32-bit OS" -ForegroundColor Green
        $program_files = "Program Files"
        $os_arch = "32"
    }

}



$createShortcut = {
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\Axia.lnk")
    $Shortcut.TargetPath = "C:\$program_files\AppLauncher.exe"
    $Shortcut.Save()
}

$set_perms =
{

    ## apply modify permission to apps folder

    write-host "applying modify permission to new PTG folder" -ForegroundColor Green

    $acl = Get-Acl -Path \\$computer\c$\$program_files\PTG\
    $perm = 'Domain Users', 'Read,Modify', 'ContainerInherit, ObjectInherit', 'None', 'Allow' 
    $rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $perm
    $acl.SetAccessRule($rule) 
    $acl | Set-Acl -Path \\$computer\c$\$program_files\PTG
} 

function installBoth() {
    #Transfer axia to client computer
 foreach ($computer in $list_of_computers) 
    {
        # Check if computer is online or not
        if (Test-Connection $computer -count 1 -Quiet)
        {
            .$check_arch
            Robocopy /s /r:1 /w:1 $axiaPath \\$computer\c$\$program_files\PTG
            Robocopy /s /r:1 /w:1 $axiaTestPath \\$computer\c$\$program_files\PTGTest
            Write-Host "Copied axia successfully at C:\$program_files\PTG and C:\$program_files\PTGTest" -ForegroundColor Green
            Robocopy /s /r:1 /w:1 $crystalPath \\$computer\c$\Temp
#			Invoke-Command -ComputerName $computer -ArgumentList $computer -ScriptBlock {
#            (Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\Temp\CRRuntime_32bit_13_0_13.msi /quiet").ExitCode
#            }
			# Create shortcut at public\public desktop for all users
            Robocopy /s /r:1 /w:1 $axiaIcon\$os_arch \\$computer\c$\Users\Public\Desktop
            Robocopy /s /r:1 /w:1 $axiaTestIcon\$os_arch \\$computer\c$\Users\Public\Desktop
            Write-Host "Created public desktop shortcut for computer at $computer" -ForegroundColor Green
        }
        else{

            # notify on screen that host is offline
            Write-Host "$computer is offline, skipping..." -ForegroundColor Yellow

        }

     }

Write-Host ""
Write-Host "Install script completed. You will still need to install the CRRuntime_32bit_13_0_13.msi." -ForegroundColor Green

}

function installProd() {
#Transfer axia to client computer
 foreach ($computer in $list_of_computers) 
    {
        # Check if computer is online or not
        if (Test-Connection $computer -count 1 -Quiet)
        {
            .$check_arch
            Robocopy /s /r:1 /w:1 $axiaPath \\$computer\c$\$program_files\PTG
            Write-Host "Copied axia successfully at C:\$program_files\PTG" -ForegroundColor Green
            Robocopy /s /r:1 /w:1 $crystalPath \\$computer\c$\Temp
#			Invoke-Command -ComputerName $computer -ArgumentList $computer -ScriptBlock {
#            (Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\Temp\CRRuntime_32bit_13_0_13.msi /quiet").ExitCode
#            }
			# Create shortcut at public\public desktop for all users
            Robocopy /s /r:1 /w:1 $axiaIcon\$os_arch \\$computer\c$\Users\Public\Desktop
            Write-Host "Created public desktop shortcut for computer at $computer" -ForegroundColor Green
        }
        else{

            # notify on screen that host is offline
            Write-Host "$computer is offline, skipping..." -ForegroundColor Yellow

        }

     }

Write-Host ""
Write-Host "Install script completed. You will still need to install the CRRuntime_32bit_13_0_13.msi." -ForegroundColor Green
}

function installTest() {
#Transfer axia to client computer
 foreach ($computer in $list_of_computers) 
    {
        # Check if computer is online or not
        if (Test-Connection $computer -count 1 -Quiet)
        {
            .$check_arch
            Robocopy /s /r:1 /w:1 $axiaTestPath \\$computer\c$\$program_files\PTGTest
            Write-Host "Copied axia successfully at C:\$program_files\PTGTest" -ForegroundColor Green
            Robocopy /s /r:1 /w:1 $crystalPath \\$computer\c$\Temp
#			Invoke-Command -ComputerName $computer -ArgumentList $computer -ScriptBlock {
#            (Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\Temp\CRRuntime_32bit_13_0_13.msi /quiet").ExitCode
#            }
			# Create shortcut at public\public desktop for all users
            Robocopy /s /r:1 /w:1 $axiaTestIcon\$os_arch \\$computer\c$\Users\Public\Desktop
            Write-Host "Created public desktop shortcut for computer at $computer" -ForegroundColor Green
        }
        else{

            # notify on screen that host is offline
            Write-Host "$computer is offline, skipping..." -ForegroundColor Yellow

        }

     }

Write-Host ""
Write-Host "Install script completed. You will still need to install the CRRuntime_32bit_13_0_13.msi." -ForegroundColor Green
}

function Show-Menu
{
     param (
           [string]$Title = 'Remotely Install Axia'
     )
     Write-Host "================ $Title ================"
    
     Write-Host "1: Install Axia Prod and Axia Test: Press '1' for this option."
     Write-Host "2: Install Axia Prod: Press '2' for this option."
     Write-Host "3: Install Axia Test '3' for this option."
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
                installBoth
           } '2' {
                cls
                'You chose option #2'
                installProd
           } '3' {
                cls
                'You chose option #3'
                installTest
           } 'q' {
                return
           }
     }
     pause
}
until ($input -eq 'q')
