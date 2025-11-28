#
# Note: when using windows task scheduler to execute the python script via a scheduled
# task, you must select the option "Run only when user is logged in".  This requirement
# is due to the fact that the python script needs to call / dispatch the outlook client
# in order to retrieve its calendar events. The outlook client requires a desktop
# environment and you only get a desktop environment for a logged in user.
#
Set-Location -Path $PSScriptRoot
.\.venv\Scripts\Activate.ps1
.\.venv\Scripts\python.exe .\cal2plannerpdf.py --autostart
