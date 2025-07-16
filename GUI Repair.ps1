#Verifying administrator permissions are granted, if not, opens in elevated window
try{
    if(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`"  `"$($MyInvocation.MyCommand.UnboundArguments)`""
    }
} catch {
    Write-Host "Unable to elevate to local administrator. You must be able to elevate to Local Administrator rights in order to run this script." -ForegroundColor Red
    Start-Sleep -Seconds 5
    return
}

#if this works, need to close other powershell window and start the logging process
$currentProcessId = $PID
$powershellProcesses = Get-Process -Name powershell
$powershellProcesses | Where-Object {$_.Id -ne $currentProcessId} | Stop-Process -Force

Start-Transcript -Path "C:\temp\LutronGUIRepair.log"

Write-Host "Elevation Succesful" -ForegroundColor Green
Start-Sleep -Seconds 5

#Initialzing Functions for Script

#Function to get response with y/n prompt
function Get-YNResponse([string]$Prompt) {
    $validInput = $false
    while (-not $validInput) {
       $response = Read-Host $Prompt
       if ($response -eq "y" -or $response -eq "n") {
           $validInput = $true
           return $response
       } else {
           Write-Host "Invalid input. Please enter 'y' or 'n'." -ForegroundColor Red
       }
   }
}

#Simple function to restart the localdb instances
function Restart-Instances {
    sqllocaldb stop mssqllocaldb
    sqllocaldb delete mssqllocaldb
    sqllocaldb create mssqllocaldb
    sqllocaldb start mssqllocaldb

    sqllocaldb stop v11.0
    sqllocaldb delete v11.0
    sqllocaldb create v11.0 11.0
    sqllocaldb start v11.0
}

#function to test if v11 instance is running, should be run after each restart instances function along with mssql one, returns boolean
function Check-v11 {
    $status = sqllocaldb info "v11.0"
    if($status -match "running")
    {
        Write-Host "v11 is running" -ForegroundColor Green
        return $true
    } elseif ($status -match "stopped") {
        Write-Host "v11 is stopped" -ForegroundColor Red
        return $false
    } else{
        Write-Host "v11 in unknown state" -ForegroundColor Red
        return $false
    }
}

#same as above but for mssqllocaldb instance
function Check-mssql {
    $status = sqllocaldb info "mssqllocaldb"
    if($status -match "running")
    {
        Write-Host "mssqllocaldb is running" -ForegroundColor Green
        return $true
    } elseif ($status -match "stopped") {
        Write-Host "mssqllocaldb is stopped" -ForegroundColor Red
        return $false
    } else{
        Write-Host "mssqllocaldb in unknown state" -ForegroundColor Red
        return $false
    }
}

#Reg edits for Sector Size fix
function Fix-SectorSize {
    Write-Host "Correcting Registry..." -ForegroundColor Green
    REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL14E.LOCALDB\MSSQLServer\Parameters" /v "SQLArg0" /t REG_SZ /d "-T1800" /f /reg:64 

    REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL11E.LOCALDB\MSSQLServer\Parameters" /v "SQLArg0" /t REG_SZ /d "-T1800" /f /reg:64 

    REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Lutron\Lutron Designer\SQLParameterUpdate" /v "IsParameterAdded" /t REG_DWORD /d 1 /f 

    Write-Host "You must restart your machine for the changes to take place, please restart your computer and re-run this script." -ForegroundColor Green
    Start-Sleep -Seconds 10
    exit
}

#function for checking sector size relying on user input
function Check-SectorSize {
    Write-Host "Verifying Sector Size..." -ForegroundColor Green
    fsutil fsinfo sectorinfo C: | findstr PhysicalBytesPerSector
    $sectorResponse = Get-YNResponse -Prompt "Are any of the above values above 4096(y/n)?"
    if($sectorResponse -eq "y"){
        Fix-SectorSize
    } elseif($sectorResponse -eq "n") {
        Write-Host "Sector Size is correct" -ForegroundColor Green
    }
}

#function for installing SSMS from microsoft
function Install-SSMS {
    Write-Host "Moving on to Install SQL Server Management Studio from Microsoft, in order to clean up any SQL dependencies that did not get installed correctly" -ForegroundColor Green
	Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder to download the installer to when the window pops up:" #could send this to C:\temp but if they do not have space on C drive it will fail
    Start-Sleep -Seconds 5
    $result = $folderBrowser.ShowDialog()

    if ($result -eq "OK") {
      $installerPath = $folderBrowser.SelectedPath
      Write-Host "Selected folder: $installerPath"
    } else {
      Write-Host "Folder selection cancelled. Re-run script to get back to this dialogue." -ForegroundColor Red
      Start-Sleep -Seconds 5
	  return
	}
	Push-Location -Path $installerPath
	Write-Host "Attempting to download installer from Microsoft, this may take a few minutes:" -ForegroundColor Green
	try{
        Invoke-WebRequest https://aka.ms/ssmsfullsetup -OutFile '.\SSMS-Setup-ENU.exe' -ErrorAction Stop
    }
    catch{
        Write-Warning "Error downloading file: $($_.exception.message)"
        Write-Host "Please re-run script after troubleshooting the above message." -ForegroundColor Red
		Pop-Location
        Start-Sleep -Seconds 5
		return
    }
	Write-Host "Installing SSMS..." -ForegroundColor Green
	$arguments = "/install","/norestart"

    Write-Verbose "Executing 'SSMS-Setup-ENU.exe $arguments'" -Verbose

    try{
        $result = Start-Process .\SSMS-Setup-ENU.exe -ArgumentList $arguments -PassThru -Wait
    }
    catch{
        Write-Warning $_.exception.message
		Write-Host "Please re-run script after troubleshooting the above message." -ForegroundColor Red
		Pop-Location
        Start-Sleep -Seconds 5
		return
    }

    switch -Exact($result.ExitCode){
        1603 {
            Write-Host "Reboot is required"
        }

        0 {
            Write-Host "Installation successful"
        }

        default{
            Write-Host "Installation was not successful"
        }
    }
	Write-Host "If installation was successful, please try opening SQL Server Management Studio once, then try opening Ra2/HWQS Designer and see if this allows it to run!" -ForegroundColor Green
	Write-Host "If this does not allow it to run, please let Tech Support know you ran this tool and it did not correct your issues, include the log file found in C:\temp\LutronGUIRepair.log." -ForegroundColor Green
    Start-Sleep -Seconds 5
	Pop-Location
}

#Start of actual Script
Write-Host "Going to attempt to repair SQL for the Ra2/HWQS Designer Software" -ForegroundColor Green
Start-Sleep -Seconds 2
#Restarting the instances on startup of script
Restart-Instances
#checking that all instances are running
#Write-Host "Pausing for 30 Seconds so you can manually stop instances for testing." #will comment out 
#Start-Sleep -Seconds 30 #adding a timer so that you can manually stop db instances for testing, will comment out eventually
$mssqlr = Check-mssql
$v11r = Check-v11
#if they are running, try opening the GUI
if($v11r -And $mssqlr)
{
    $initialResponse = Get-YNResponse -Prompt "The SQL tables have been succesfully restarted, please try opening the RA2/HWQS Designer software. Did this step allow it to run?(y/n)"
    #if yes, issue resolved
    if($initialResponse -eq "y"){
        Write-Host "Glad we were able to get this resolved!" -ForegroundColor Green
        Start-Sleep -Seconds 5
        return
    } elseif($initialResponse -eq "n"){ #if this did not resolve it, move onto next step
        Write-Host "Moving onto next step..." -ForegroundColor Green
    }
} else { #ONE OR BOTH INSTANCES FAILED TO START
    Write-Host "One or more of the SQL instances failed to start, moving onto next step..." -ForegroundColor Green
    #running the sector size fix
    $sectorCheck = Get-YNResponse -Prompt "Have you ran through this script before and already verified Sector Size is good?(y/n)"
    if($sectorCheck -eq "y"){
        Write-Host "Moving onto next step..." -ForegroundColor Green
    } else {
        Check-SectorSize
    }
}
#if the above did not fix it, install SSMS and try again
Write-Host "Moving onto install SSMS..." -ForegroundColor Green
Install-SSMS
Start-Sleep -Seconds 60 
Stop-Transcript

