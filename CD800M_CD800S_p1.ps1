[Console]::OutputEncoding = [System.Text.Encoding]::UTF8 # This line is added automatically to help with Unicode characters. Please add your code below 

#Please reach out to khoaphan@costco.com for any questions.

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

#----------------------------------------------------------------------------------------------------------------
#FIND AND DELETE PRINTER
function verifyPrinter {
    param([string]$printer)
    $found = $false
    $deleted = $false

    try {
    Get-Printer -Name $printer -ErrorAction Stop | Out-Null
    $found = $true
    Remove-Printer -Name $printer

    #------------------------------------------------------------------------------------------------------------
    #VERIFY PRINTER IS DELETED
        try {
            Get-Printer -Name $printer -ErrorAction Stop | Out-Null
            Write-Host "$printer was found but could not be deleted."
        } catch {
            $deleted = $true
            Write-Host "$printer was found and deleted."
        }
    #------------------------------------------------------------------------------------------------------------
    #ERROR IF PRINTER CANNOT BE FOUND
    } catch {
        #Write-Host "$printer cannot be found. Continuing..."
    }

    $printerObject = [PSCustomObject]@{
        PrinterName = $printer
        Found = $found
        Deleted = $deleted
        Date = $((Get-Date).ToString())
    }

    return $printerObject
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
#MAIN SCRIPT

Clear-Host
Write-Host '****************************************' -BackgroundColor White -ForegroundColor Black
Write-Host '******** CD800 Card Printer Fix ********' -BackgroundColor White -ForegroundColor Black
Write-Host '****************************************' -BackgroundColor White -ForegroundColor Black
Write-Host

#----------------------------------------------------------------------------------------------------------------
#FIND AND DELETE PRINTER

$cd800m = verifyPrinter -printer "CD800M"
$cd800s = verifyPrinter -printer "CD800S"
$xps = verifyPrinter -printer "XPS Card Printer"

$listOfPrinters = New-Object 'System.Collections.Generic.List[psobject]'
$listOfPrinters.Add($cd800m)
$listOfPrinters.Add($cd800s)
$listOfPrinters.Add($xps)

$listOfPrinters | Format-Table #Displays results in table format

#EXPORTS TEXT FILE TO C:\TEMP IF PRINTER WAS DETECTED BUT CANNOT BE DELETED
foreach ($printer in $listOfPrinters){
    if ($printer.Found -eq $true -and $printer.Deleted -eq $false) {
        $listOfPrinters | Format-Table | Out-file C:\Temp\CD800_Printer_Error.txt
        Write-Host "Unable to remove one or more printers, please escalate to Client Hardware Support. Details can be found in C:\Temp\CD800_Printer_Error.txt."
        controlUpSend "Unable to remove card printer from Print Management. Details can be found in C:\Temp\CD800_Printer_Error.txt."
    }
}

#----------------------------------------------------------------------------------------------------------------
#FIND AND DELETE PORT

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
        $listOfPorts | Format-Table | Out-file C:\Temp\CD800_Port_Error.txt
        Write-Host "Unable to remove one or more ports, please escalate to Client Hardware Support. Details can be found in C:\Temp\CD800_Port_Error.txt."
        controlUpSend "Unable to remove card printer port. Details can be found in C:\Temp\CD800_Port_Error.txt."
    }
}

#----------------------------------------------------------------------------------------------------------------
#FIND AND DELETE USB ROOT HUBS

$listOfDrivers = New-Object 'System.Collections.Generic.List[psobject]'

$rootHub = verifyDriver -driver "USB Root Hub"
$rootHub3 = verifyDriver -driver "USB Root Hub (USB 3.0)"

$listOfDrivers.Add($rootHub)
$listOfDrivers.Add($rootHub3)

$listOfDrivers | Format-Table #Displays results in table format

#EXPORTS TEXT FILE TO C:\TEMP IF DRIVER WAS DETECTED BUT CANNOT BE DELETED
foreach ($driver in $listOfDrivers){
    if ($driver.Found -eq $true -and $driver.Deleted -eq $false) {
        $listOfDrivers | Format-Table | Out-file C:\Temp\CD800_Driver_Error.txt
        Write-Host "Unable to remove one or more drivers, please escalate to Client Hardware Support. Details can be found in C:\Temp\CD800_Driver_Error.txt."
        controlUpSend "Unable to remove card printer driver. Details can be found in C:\Temp\CD800_Driver_Error.txt."
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
        $listOfNet | Format-Table | Out-file C:\Temp\CD800_NetAdapter_Error.txt
        Write-Host "Unable to remove one or more drivers, please escalate to Client Hardware Support. Details can be found in C:\Temp\CD800_NetAdapter_Error.txt."
        controlUpSend "Unable to remove card printer driver. Details can be found in C:\Temp\CD800_NetAdapter_Error.txt."
    }
}

#----------------------------------------------------------------------------------------------------------------
#OUTPUTS DATA TO LOCAL EVENT VIEWER (FOR CONTROLUP ONLY)

controlUpSend -output "CD800M Fix - Part 1 ran."

#----------------------------------------------------------------------------------------------------------------
#REBOOTS
Write-Host "Computer is rebooting now..."
controlUpSend -output "Computer is now rebooting."
shutdown /r /t 0