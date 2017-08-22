#-----------------------------------------------------------------------------
# Script: disableTouchScreen
# Author: Zach Brisson
# Date: 08/22/17
# Prerequisites: Devcon must be installed(preferably in C:\Windows\System32)
# Searches for the Touch Screen Driver and disables it. 
# Once we move away from Desktop Authority for drive mappings this script can be deleted.


$id = (Get-CimInstance Win32_PnPEntity |

where caption -match 'touch screen').pnpDeviceID

$ppid = "{0}{1}" -f '@',$id

Devcon disable $ppid
