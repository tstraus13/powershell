###################################################################################################
# 
# Create by Tom Strausbaugh 
# 2/14/2018
#
# This is a script to copy an installation file(s) to a computer
# and run the installation remotely. It was created for installing
# McAfee because they did not readily supply an MSI installer for their
# agent. It could be used for other installers as well that allow for 
# silent/quite install. Installation files are removed when finished.
#
# @param installFileSource (required) The location of the install file(s). Can be a directory.
#   Whether there is one install file or many there needs to be an associate text file with
#   switches for silent/quiet install. (Example: vlc-install.exe & vlc-install.exe.txt)
#
# @param computersFile (optional) The location of the text file with list of computers.
#   Default is computers.txt located wherever you exectuted this script.
#
# @param organizationalUnit (optional) Which OU you want to use to pull a list of computers
#
# @param checkProcessWaitTime (optional) The length of time, in seconds, to wait between checking
# if the current installation is still running.
#   Default is 30 seconds
#
###################################################################################################

param(
    [Parameter(Mandatory=$TRUE)][string]$installFileSource,
    [string]$computersFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\computers.txt",
    [string]$organizationalUnit,
    [int]$checkProcessWaitTime = 30
)

Import-Module ActiveDirectory


# Checks if installFileSource is a directory or not. Sets the install file name(s)
# accordingly based on if the source is a folder or a single file.
if ((Get-Item "$installFileSource") -is [System.IO.DirectoryInfo])
{
    $installFileNames = Split-Path -Path "$installFileSource\*.exe" -Leaf -Resolve
    $isFolder = $TRUE
}

else 
{
    $installFileNames = Split-Path -Path "$installFileSource" -Leaf -Resolve
    $isFolder = $FALSE
}


# Check to make sure there is install file(s) at given location
if(!$installFileNames)
{
    Write-Host "No install files found. Exiting..."
    exit
}


# Checks if an OU is provided and if so pull all the computers/machines from that OU
# and confirm with user that the list is good to use. Otherwise use the computer list 
# text file.
if ($organizationalUnit)
{
    $computers = Get-ADComputer -Filter * -SearchBase "OU=$organizationalUnit, DC=tsg, DC=piedmont-airlines, DC=com" | Select -Expand Name
    
    foreach ($computer in $computers)
    {
        Write-Host "$computer"
    }
    
    $isListOK = Read-Host "Does this list of computer look alright (y / n)?"

    Switch ($isListOK)
    {
        Y { }
        N { Write-Host "Exiting..."; exit }
        Default { Write-Host "Exiting..."; exit }
    }
}

else
{
    # Open computer list file and save to array
    $computers = Get-Content "$computersFile"
}


# Loop through every computer in the array
foreach($computer in $computers)
{
    # Check to see if folder where we will copy install file exists.
    # If it does not exist, create it.
    if (!(Test-Path -Path "\\$computer\c$\InstallFiles"))
    {
        $installFolder = New-Item -ItemType directory -Path "\\$computer\c$\InstallFiles"
    }

    foreach ($install in $installFileNames)
    {
        Write-Host "Starting install of $install to $computer. Please Wait..."

        # Copy installation file to remote computer in the install files directory
        # and get the install file switches from associated text file.
        if ($isFolder)
        {
            Copy-Item "$installFileSource\$install" -Destination "\\$computer\c$\InstallFiles\$install"
            $installFileSwitches = Get-Content "$installFileSource\$install.txt"
        }

        else
        {
            Copy-Item "$installFileSource" -Destination "\\$computer\c$\InstallFiles\$install"
            $installFileSwitches = Get-Content "$installFileSource.txt"
        }

        # Spawn a new process and run the install file with switches
        $installProcess = ([WMICLASS]"\\$computer\ROOT\CIMV2:Win32_Process").Create("C:\InstallFiles\$install $installFileSwitches")

        # Check to see if installation is still running. If so, wait 30 seconds and check again
        while (Get-WmiObject Win32_Process -computername "$computer" -filter "ProcessID='$($installProcess.ProcessID)'")
        {
            Start-Sleep -seconds $checkProcessWaitTime
            $datetime = Get-Date -UFormat "%Y-%m-%d %r"
            Write-Host "$datetime - $install installation is still running on $computer. Waiting $checkProcessWaitTime Seconds..."
        }

        Write-Host "Finished installation of $install on $computer. Moving onto next installation file."
    }

    # Delete install folder and any installation files from remote computer
    Remove-Item -Path "\\$computer\c$\InstallFiles" -Force -Recurse

    Write-Host "Finished installation(s) and Deleted all installation files on $computer. Moving onto next computer..."
}

Write-Host "Finished installation on all provided computers! Exiting..."