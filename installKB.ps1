<#
.SYNOPSIS
installKB.ps1 - Windows Update installer Script

.DESCRIPTION 
A PowerShell script to install Windows updates.

This script will check if the named KB is installed on the device,
and if not, attempt to install the update with the provided .wsu file.

The recommended use for this script is to install updates that or out-of-band.
Parameters allow the script to be executed with different KBs without having to
modify the script itself.

If the script detects any issues, it will not attempt to install the update.

.PARAMETER Name
The name of the KB to detect installation status on the endpoint.

.PARAMETER Path
The path of the Windows update to be installed if the KB is not detected.

.PARAMETER ForcedRestart
Switch parameter to force a restart or prompt for a restart after a succeddful install.

.EXAMPLE
.\installKB.ps1 -Name "KB5004946" -Path "D:\updates\windows10.0-kb5004946-x64.msu" -ForceRestart
Thsi example will check for the installed KB5004946, install the update specified in Path in the KB
is not detected, and will force the edpoint to restart once update is successfully installed.

.EXAMPLE
.\installKB.ps1 -Name "KB5004946" -Path "D:\updates\windows10.0-kb5004946-x64.msu"
Thsi example will check for the installed KB5004946, install the update specified in Path in the KB
is not detected, and will prompt the user to reboot once the update is successfully installed.

.LINK


.NOTES
Written by: Jason Thweatt

License:

The MIT License (MIT)

Copyright (c) 2021 Jason Thweatt

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Change Log
V1.00, 2021-07-09, Initial build.
#>


[CmdletBinding()]
param (
    # Name of KB to detect installed update.
    [Parameter(Mandatory=$true)]
    [string]
    $Name,

    # File path of update to install.
    [Parameter(Mandatory=$true)]
    [string]
    $Path,

    # Force restart on successful install.
    [Parameter(Mandatory=$false)]
    [switch]
    $ForceRestart
)

# Detect if endpoint is online.
Write-Host "Testing connection to $env:ComputerName..."
if(Test-Connection -ComputerName $env:ComputerName -Quiet)
{
    Write-Host "$env:ComputerName is online." -ForegroundColor Green

    # Detect valid file path for installer.
    Write-Host "Testing installer file path..."
    if(Test-Path $Path){
        Write-Host "Installer path is valid." -ForegroundColor Green

        # Detect if KB is installed and attempt install if not.
        Write-Host "Checking for $Name..."
        if(get-hotfix | where-object {$_.HotFixId -match $Name}){
            Write-Host "$Name is already install on $env:ComputerName." -ForegroundColor Green
        }
        else{
            Write-Host "$Name is not install on $env:ComputerName." -ForegroundColor Yellow

            # Stop Windows Update Service if running.
            Write-Host "Stopping Windows Update Service..."

            Get-Service -DisplayName "Windows Update" -ComputerName $env:ComputerName | Stop-Service -Force

            # Check if service stopped successfully
            $service =  Get-Service -DisplayName "Windows Update" -ComputerName $env:ComputerName
            if ($service.status -ne "Stopped"){
                Write-Host "Windows Update service did not stop successfully. Try running the script elevated." -ForegroundColor Yellow
            }
            else {
                Write-Host "Windows Update service stopped successfully." -ForegroundColor Green
            }

            # Install KB
            Write-Host "Installing $Name..."

            if (!(Test-Path $env:systemroot\SysWOW64\wusa.exe))
            {
                $wusa = "$env:systemroot\System32\wusa.exe"
            }
            else 
            {
                $wusa = "$env:systemroot\SysWOW64\wusa.exe"
            }

            Start-Process -FilePath $wusa -argumentlist "$Path /quiet /norestart" -Wait
            
            if (get-hotfix | where-object {$_.HotFixId -match $Name}) {
                Write-Host "$Name has successfully installed on $env:ComputerName." -ForegroundColor Green

                if($ForceRestart){
                    Write-Warning "THE DEVICE WILL NOW RESTART."
                    Start-Sleep -Seconds 30
                    Restart-Computer -Force
                }
                else {
                    Write-Warning "DEVICE RESTART PENDING."
                    Restart-Computer -Confirm
                }
            }
            else{
                Write-Host "$Name failed to install on $env:ComputerName. Check endpoint." -ForegroundColor Red
            }
        }
    }
    else {
        Write-Host "File path is invalid. Verify installer path." -ForegroundColor Red
    }
}
else {
    Write-Host "Endpoint $env:ComputerName in not online. Is the device on?" -ForegroundColor Red
}
