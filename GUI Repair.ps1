Start-Transcript -Path "C:\temp\LutronGUIRepair.log"
Write-Host "Log file saved at: C:\temp\LutronGUIRepair.log"
#Admin level check, if fsutil is empty, they did not run as admin
$adminTest = fsutil fsinfo sectorinfo C: | findstr PhysicalBytesPerSector
if($adminTest.Length -eq 0)
{
    Write-Host "This application will close, please right click the .exe and run as admin to proceed." -ForegroundColor Red
    Start-Sleep -Seconds 60
    exit
}

#Write-Host "Elevation Succesful" -ForegroundColor Green
Start-Sleep -Seconds 5

#Initialzing Functions for Script

#Function to get response with y/n prompt
function Get-YNResponse([string]$Prompt) {
    $validInput = $false
    while (-not $validInput) {
        while ([Console]::KeyAvailable) { #Clears input stream so only most recent press after this will be used as input
            [Console]::ReadKey($true) 
        }
        Write-Host $Prompt -ForegroundColor Green
       $response = [System.Console]::ReadKey($true).KeyChar.toString()
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
    REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL14E.LOCALDB\MSSQLServer\Parameters" /v "SQLArg0" /t REG_SZ /d "-T1800" /f /reg:64 | Out-Null

    REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL11E.LOCALDB\MSSQLServer\Parameters" /v "SQLArg0" /t REG_SZ /d "-T1800" /f /reg:64  | Out-Null

    REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Lutron\Lutron Designer\SQLParameterUpdate" /v "IsParameterAdded" /t REG_DWORD /d 1 /f  | Out-Null
    
    REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\stornvme\Parameters\Device" /v "ForcedPhysicalSectorSizeInBytes" /t REG_MULTI_SZ /d "* 4095" /f /reg:64 | Out-Null

    Write-Host "You must restart your machine for the changes to take place, please restart your computer and re-run this program." -ForegroundColor Green
    Start-Sleep -Seconds 10
}

#function for checking sector size relying on user input
function Check-SectorSize {
    Write-Host "Verifying Sector Size..." -ForegroundColor Green
    $fsutilOutput = fsutil fsinfo sectorinfo C: | findstr PhysicalBytesPerSector
    $SectorValues = $fsutilOutput -split "\D+"
    $sectorGood = $true
    foreach($value in $SectorValues) {
        if($value -match "\d+") {
            if($value -gt 4096)
            {
                Write-Host "Value is $value"
                $sectorGood = $false
            }
         }
    }

    if($sectorGood -eq $false){
        Fix-SectorSize
        return "Fix Ran"
    } elseif($sectorGood -eq $true) {
        Write-Host "Sector Size is correct" -ForegroundColor Green
        return "Sector Size OK"
    }
}

#function for installing SSMS from microsoft
function Install-SSMS {
    $installerPath = "C:\temp"
	Push-Location -Path $installerPath
	Write-Host "Attempting to download installer from Microsoft, this may take a few minutes:" -ForegroundColor Green
	try{
        Invoke-WebRequest https://aka.ms/ssmsfullsetup -OutFile '.\SSMS-Setup-ENU.exe' -ErrorAction Stop
    }
    catch{
        Write-Warning "Error downloading file: $($_.exception.message)"
        Write-Host "Please re-run program after troubleshooting the above message or alternatively downloading and installing SQL Server Management Studio from Microsoft manually." -ForegroundColor Red
		Pop-Location
        Start-Sleep -Seconds 5
		return
    }
	Write-Host "Installing SSMS..." -ForegroundColor Green
	$arguments = "/install","/norestart"

    Write-Host "Executing 'SSMS-Setup-ENU.exe $arguments'" -ForegroundColor Green

    try{
        $result = Start-Process .\SSMS-Setup-ENU.exe -ArgumentList $arguments -PassThru -Wait
    }
    catch{
        Write-Warning $_.exception.message
		Write-Host "Please re-run program after troubleshooting the above message." -ForegroundColor Red
		Pop-Location
        Start-Sleep -Seconds 5
		exit
    }
    $exitCodeDisplay = $result.ExitCode 
    Write-Host "This is SSMS exit code $exitCodeDisplay" -ForegroundColor Red
    switch -Exact($exitCodeDisplay){
        1603 { #this is fatal so case should not occur since it should catch generally
            Write-Host "Reboot is required"
        }

        0 {
            Write-Host "Installation successful"
        }
        
        3010 {
            Write-Host "Reboot is required"
        }

        1626 {
            Write-Host "Looks like there is a pending restart, please restart your machine and try again" -ForegroundColor Red
        }

        default{
            Write-Host "Installation was not successful" -ForegroundColor Red
            Write-Host "Please re-run script after trying to manually install SQL Server Management Studio from Microsoft's Website." -ForegroundColor Red
            Pop-Location
            Start-Sleep -Seconds 10
            exit
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
Start-Sleep -Seconds 10 #adding a timer so that you can manually stop db instances for testing, will comment out eventually
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
    $sectorSizeFile = "C:\temp\SSF.txt"
    if(Test-Path -Path $sectorSizeFile -PathType Leaf) {
        #File exists, meaning they have been through this step before
        $sectorResult = Get-Content $sectorSizeFile
        if($sectorResult -eq "Fix Ran") { #this step has ran before and ran the regedits fix

           $rebootCheck = Get-YNResponse "Have you rebooted your machine since running the last step(y/n)"

           if($rebootCheck -eq "y") { #if yes, verify that the change took effect
                Write-Host "Verifying Sector Size..." -ForegroundColor Green
                $fsutilOutput = fsutil fsinfo sectorinfo C: | findstr PhysicalBytesPerSector
                $SectorValues = $fsutilOutput -split "\D+"
                $sectorGood = $true
                foreach($value in $SectorValues) {
                    if($value -match "\d+") {
                        if($value -gt 4096) {
                            $sectorGood = $false
                        }
                    }
                }
                if($sectorGood -eq $false){ #rebooted and sector size changes did not take effect
                    Write-Host "The sector size changes did not work as expected. Please reach out to Lutron Tech Support and let them know you ran through this program. Include the files in the C:\temp directory." -ForegroundColor Red
                    Start-Sleep -Seconds 900 #sleep until they close out
                    exit
                } elseif($sectorResponse2 -eq $true){ #rebooted and sector size changes did take effect, but instances still did not start
                    Write-Host "Sector Size changes worked, but the instances still failed to start, moving to next step..."
                    Start-Sleep -Seconds 5
                }
           } elseif($rebootCheck -eq "n") { #meaning they ran the sector size fix but have not rebooted computer
                Write-Host "You must restart your machine for the Sector Size changes to take effect." -ForegroundColor Red
                Start-Sleep -Seconds 10
                exit
           }
        } elseif($sectorResult -eq "Sector Size OK"){
            Write-Host "Sector Size is ok, but the instances still failed to start, moving to next step..."
            Start-Sleep -Seconds 5
        }
    } else {
        #File does not exist, they have not been through this step before
        $sectorSizeResult = Check-SectorSize
        Set-Content -Path $sectorSizeFile -Value $sectorSizeResult
        if($sectorSizeResult -eq "Fix Ran"){ #if the fix was ran, let it save to the file and exit, otherwise keep going
            exit
        }
    }
}
#if the above did not fix it, install SSMS and try again
Write-Host "Moving onto install SSMS..." -ForegroundColor Green
Start-Sleep -Seconds 10
Install-SSMS
Start-Sleep -Seconds 900 
Stop-Transcript