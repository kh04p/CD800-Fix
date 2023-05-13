[Console]::OutputEncoding = [System.Text.Encoding]::UTF8 # This line is added automatically to help with Unicode characters. Please add your code below 

#Please reach out to khoaphan@costco.com for any questions.

#----------------------------------------------------------------------------------------------------------------
#SET PRINTER AS DEFAULT
function setDefault {
    param([string]$printerName)
    $printer = Get-CimInstance -Class Win32_Printer -Filter "Name='$printerName'"
    Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter | Out-Null
}

#----------------------------------------------------------------------------------------------------------------
#FIND AND DELETE PORT 
function verifyPort {
    param([string]$printerPort)
    $found = $false
    $deleted = $false

    try {
        Get-PrinterPort -Name $printerPort -ErrorAction Stop | Out-Null
        $found = $true
        

        #------------------------------------------------------------------------------------------------------------
        #VERIFY PORT IS DELETED
            try {
                Remove-PrinterPort -Name $printerPort -ErrorAction Stop | Out-Null
                $deleted = $true
                Write-Host "$printerPort was found and deleted."
            } catch {
                Write-Host "$printerPort was found but could not be deleted."                
            }
    #------------------------------------------------------------------------------------------------------------
    #ERROR IF PORT CANNOT BE FOUND
    } catch {
        #Write-Host "$printerPort cannot be found. Continuing..."
    }

    $portObject = [PSCustomObject]@{
        PrinterPort = $printerPort
        Found = $found
        Deleted = $deleted
        Date = $((Get-Date).ToString())
    }

    return $portObject
}

#----------------------------------------------------------------------------------------------------------------
#FIND AND DELETE DRIVERS
function verifyDriver {
    param([string]$driver)
    $found = $false
    $deleted = $false

    #------------------------------------------------------------------------------------------------------------
    #CHECK IF DRIVER EXISTS
    $verify = Get-PnpDevice | Where-Object{$_.Name -eq $driver}
    if ($verify -eq $null -or $verify -eq "") {
        #Write-Host "Device does not exist."
        $found = $false        
    } else {
        #Write-Host "Device found."
        $found = $true

        foreach ($dev in (Get-PnpDevice | Where-Object{$_.Name -eq $driver})) {
            $temp = &"pnputil" /remove-device $dev.InstanceId
        }

        $verify = Get-PnpDevice | Where-Object{$_.Name -eq $driver}
        if ($verify -eq $null -or $verify -eq "") {
            #Write-Host "Device deleted."
            $deleted = $true
        } else {
            #Write-Host "Device still exists."
            $deleted = $false
        }
    }    

    $driverObject = [PSCustomObject]@{
            Driver = $driver 
            Found = $found
            Deleted = $deleted
            Date = $((Get-Date).ToString())
    }

    return $driverObject
}

#----------------------------------------------------------------------------------------------------------------
#SENDS LOG DATA TO CONTROLUP
function controlUpSend {
    param([string]$output)

    #Gets current date
    [string]$strDate = Get-Date -Format 'dd-MM-yyyy'

    #Creates custom object to store info
    $resultObject = [PSCustomObject]@{
	    'Date'    = $strDate
	    'Event'   = $output
    }

    #Creates custom string to log into Event Viewer
    $date = (Get-Date).ToString()
    $strObject = "On $date, $output"

    #Outputs data to custom DX index card_printer. Can be accessed in Data section.
    Write-Output("### SIP DATA BEGINS ###")
    Write-Output -InputObject $resultObject | ConvertTo-Json
    Write-Output("### SIP DATA ENDS ###")

    #Outputs data to local Event Viewer
    Write-Output("### SIP EVENT BEGINS ###")
    Write-Output -InputObject $strObject
    Write-Output("### SIP EVENT ENDS ###")
}

#----------------------------------------------------------------------------------------------------------------\
#CHECKS DEVICE MANAGER
function queryDeviceManager {
    param([string]$name)
    try {
        Get-PnpDevice | Where-Object{$_.Name -eq $name} -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Host "Unable to locate ""$name""."
        return $false
    }
}

#----------------------------------------------------------------------------------------------------------------
#CHECKS PRINT MANAGEMENT
function queryPrinter {
    param([string]$printer)
    try {
        Get-Printer -Name $printer -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

#----------------------------------------------------------------------------------------------------------------
#RENAMES PRINTER
function renamePrinter {
    param([string]$wrongName)
    $printer = "CD800S"

    #Scans device manager for hardware changes, should add USB DXP01 port
    pnputil.exe /scan-devices
    #Check if printer name is already CD800S
    if ($wrongName -eq $printer) {
        #Sets printer as default
        setDefault -printerName $printer
        return $true
    }

    #If name is NOT CD800S, renames to CD800S
    Rename-Printer -Name $wrongName -NewName $printer

    #Queries to see if printer exists now
    $confirmExist = queryPrinter -printer $printer
    if ($confirmExist -eq $true) {
        #Sets default if printer exists
        setDefault -printerName $printer
        return $true
    } else {
        #Return fail code if printer cannot be detected under new name
        return $false         
    }
}

#----------------------------------------------------------------------------------------------------------------
#ADDS PRINTER IF IT DOES NOT EXIST
function addPrinter {
    $printer = "CD800S"

    #Scans device manager for hardware changes, should add USB DXP01 port
    pnputil.exe /scan-devices

    #Waits 5 seconds for net adapter to auto install
    Start-Sleep -Seconds 5

    #Adds CD800S Printer
    Add-Printer -Name "CD800S" -DriverName "XPS Card Printer" -PortName "USB DXP01 Port"
 
    #Queries to see if printer exists now
    $confirmExist = queryPrinter -printer $printer
    if ($confirmExist -eq $true) {
        #Sets default if printer exists
        setDefault -printerName $printer
        return $true
    } else {
        #Return fail code if printer cannot be detected under new name
        return $false
            
    }
}

#----------------------------------------------------------------------------------------------------------------
#MAIN SCRIPT

Clear-Host
Write-Host '****************************************' -BackgroundColor White -ForegroundColor Black
Write-Host '******* CD800 Card Printer Fix *********' -BackgroundColor White -ForegroundColor Black
Write-Host '****************************************' -BackgroundColor White -ForegroundColor Black

$name1 = "XPS Card Printer"
$name2 = "CD800"
$name3 = "CD800S"
$repeatCounter = 0
$portError = $false
$driverError = $false

#Checks if script has repeated itself more than twice
while ($repeatCounter -le 1) {

    #------------------------------------------------------------------------------------------------------------
    #CHECKS DEVICE MANAGER
    $confirmDevice = queryDeviceManager -name "XPS Card Printer High Speed USB Connection"
    if ($confirmDevice -eq $false) {
        Write-Warning "Card printer cannot be found in Device Manager. Please troubleshoot using KB2016304 and escalate if needed."
        controlUpSend -output "Card Printer driver cannot be detected."
    }

    #------------------------------------------------------------------------------------------------------------
    #CHECKS PRINTER PORT
    $usb1 = verifyPort -printerPort "USB DXP01 Port"
    $usb2 = verifyPort -printerPort "USB DXP01 Port (Copy 1)"
    $usb3 = verifyPort -printerPort "USB DXP01 Port (Copy 2)"

    $listOfPorts = New-Object 'System.Collections.Generic.List[psobject]'
    $listOfPorts.Add($usb1)
    $listOfPorts.Add($usb2)
    $listOfPorts.Add($usb3)

    $listOfPorts | Format-Table #Displays results in table format

    #EXPORTS TEXT FILE TO C:\TEMP IF PORT WAS DETECTED BUT CANNOT BE DELETED
    foreach ($port in $listOfPorts){
        if ($port.Found -eq $true -and $port.Deleted -eq $false) {
            $listOfPorts | Format-Table | Out-file C:\Temp\CD800S_Port_Error.txt
            #Write-Host "Unable to remove port $name, please escalate to Client Hardware Support. Details can be found in C:\Temp\CD800_Port_Error.txt."
            #controlUpSend "Unable to remove card printer port $name, details can be found in C:\Temp\CD800S_Port_Error.txt."
            $portError = $true
        }
    }
    
    #----------------------------------------------------------------------------------------------------------------
    #FIND AND DELETE NET ADAPTER

    $listOfNet = New-Object 'System.Collections.Generic.List[psobject]'

    $netAdapter = verifyDriver -driver "XPS Card Printer High Speed USB Connection"

    $listOfNet.Add($netAdapter)

    $listOfNet | Format-Table #Displays results in table format

    #EXPORTS TEXT FILE TO C:\TEMP IF DRIVER WAS DETECTED BUT CANNOT BE DELETED
    foreach ($driver in $listOfNet){
        if ($driver.Found -eq $true -and $driver.Deleted -eq $false) {
            $listOfNet | Format-Table | Out-file C:\Temp\CD800S_NetAdapter_Error.txt
            #Write-Host "Unable to remove one or more drivers, please escalate to Client Hardware Support. Details can be found in C:\Temp\CD800_NetAdapter_Error.txt."
            #controlUpSend "Unable to remove card printer driver. Details can be found in C:\Temp\CD800_NetAdapter_Error.txt."
            $driverError = $true
        }
    }

    #------------------------------------------------------------------------------------------------------------
    #CHECKS PRINT MANAGEMENT

    #Checks and renames XPS Card Printer
    $confirmPrinter = queryPrinter -printer $name1
    Write-Host "Printer ""$name1"" found: $confirmPrinter"
    if ($confirmPrinter -eq $true) {
        $rename = renamePrinter -wrongName $name1
        if ($rename -eq $true) {
            Write-Host "Please do a test print now. If issues persist, please troubleshoot using KB2016304 and escalate if needed."
            controlUpSend -output "Script completed successfully, please do a test print now. If issues persist, please troubleshoot using KB2016304 and escalate if needed."
            exit
        } else {
            Write-Warning "Printer is no longer detected after renaming, attempting to rerun script..."
            controlUpSend -output "Card printer cannot be detected after renaming, script will attempt to run again."
            $repeatCounter += 1
            continue
        }
    } else {

        #--------------------------------------------------------------------------------------------------------
        #Checks and renames CD800
        $confirmPrinter = queryPrinter -printer $name2
        Write-Host "Printer ""$name2"" found: $confirmPrinter"
        if ($confirmPrinter -eq $true) {

            #----------------------------------------------------------------------------------------------------
            #Renames and checks again
            $rename = renamePrinter -wrongName $name2
            if ($rename -eq $true) {
                Write-Host "Please do a test print now. If issues persist, please troubleshoot using KB2016304 and escalate if needed."
                controlUpSend -output "Script completed successfully, please do a test print now. If issues persist, please troubleshoot using KB2016304 and escalate if needed."
                exit
            } else {
                Write-Warning "Printer is no longer detected after renaming, attempting to rerun script..."
                controlUpSend -output "Card printer cannot be detected after renaming, script will attempt to run again."
                $repeatCounter += 1
                continue
            }
        } else {

            #----------------------------------------------------------------------------------------------------
            #Checks and renames CD800S
            $confirmPrinter = queryPrinter -printer $name3
            Write-Host "Printer ""$name3"" found: $confirmPrinter"
            if ($confirmPrinter -eq $true) {
            
                #------------------------------------------------------------------------------------------------
                #Renames and check again
                $rename = renamePrinter -wrongName $name3
                if ($rename -eq $true) {
                    Write-Host "Please do a test print now. If issues persist, please troubleshoot using KB2016304 and escalate if needed."
                    controlUpSend -output "Script completed successfully, please do a test print now. If issues persist, please troubleshoot using KB2016304 and escalate if needed."
                    exit
                } else {
                    Write-Warning "Printer is no longer detected after renaming, attempting to rerun script..."
                    controlUpSend -output "Card printer cannot be detected after renaming, script will attempt to run again."
                    $repeatCounter += 1
                    continue
                }
            } else {

                #----------------------------------------------------------------------------------------
                #If no printers are detected
                Write-Warning "No printers detected, printer will now be added."
                $confirmPrinter = addPrinter
                if ($confirmPrinter -eq $true) {
                    Write-Host "Please do a test print now. If issues persist, please troubleshoot using KB2016304 and escalate if needed."
                    controlUpSend -output "Card printer was added successfully, please do a test print now. If issues persist, please troubleshoot using KB2016304 and escalate if needed."
                    exit
                } else {
                    if ($portError -eq $true) {
                        Write-Warning "Printer cannot be added, possibly due to port not being removed. Please reboot computer and try again."
                        controlUpSend -output "Printer cannot be added, possibly due to port not being removed. Please reboot computer and try again. Logs are located in C:\Temp."
                        exit
                    }
                    if ($driverError -eq $true) {
                        Write-Warning "Printer cannot be added, possibly due to driver not being removed. Please reboot computer and try again."
                        controlUpSend -output "Printer cannot be added, possibly due to driver not being removed. Please reboot computer and try again. Logs are located in C:\Temp."
                        exit
                    } else {
                        Write-Warning "Printer cannot be added, please troubleshoot using KB2016304 and escalate if needed."
                        controlUpSend -output "Card printer could not be added for unknown reason, please troubleshoot using KB2016304 and escalate if needed. Logs are located in C:\Temp."
                        exit
                    }                    
                }
            }
        }
    }
}

#----------------------------------------------------------------------------------------------------------------
#EXITS IF SCRIPT HAS REPEATED OVER 2 TIMES
Write-Warning "Script exceeded maximum amount of run cycles. Please troubleshoot using KB2016304 and escalate if needed."
controlUpSend -output "Escalation needed, script failed to resolve issue."
exit