#Created by Brad Johnson
#Created 2/19/2019

#Purpose: Install all HQ printers
#Source file for printers to install: \\<NETWORK SHARE>\<PRINTERS>.csv

#Capture destination computer
$ComputerName = Read-Host "Enter computer name (or leave blank for localhost)"

#If the name is blank, use localhost
$InstallingLocally = $FALSE
if ($ComputerName -eq ""){
    $ComputerName = $env:COMPUTERNAME
    $InstallingLocally = $TRUE
}

#Import CSV file
$Printers = Import-Csv "\\<NETWORK SHARE>\<PRINTERS>.csv"

#Loop through each row containing printer details in the CSV file
foreach ($Printer in $Printers) {
    #Read printer data from each field in each row and assign the data to a variable as below
    
    $PrinterName = $Printer.Name
    $PrinterIP = $Printer.IP
    $DriverName = $Printer.DriverName
    $DriverPath = $Printer.DriverPath


    # Remove printer if already installed
    Remove-Printer -ComputerName $ComputerName -Name $PRINTERNAME -ErrorAction "silentlycontinue"
    Remove-PrinterPort -ComputerName $ComputerName -Name $PRINTERIP -ErrorAction "silentlycontinue"

    # Copy and install drivers to machine
    robocopy "\\<NETWORK SHARE>\$DRIVERPATH" "\\$ComputerName\c$\temp\PRINTER DRIVERS" /mir
    if ($InstallingLocally){ #Installing locally. Run pnputil.exe directly.
        pnputil.exe /a "C:\temp\PRINTER DRIVERS\OEMSETUP.INF"
    }else{ #Installing remotely. Run pnputil.exe remotely via Invoke-Command.
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {pnputil.exe /a "C:\temp\PRINTER DRIVERS\OEMSETUP.INF"}
    }

    # Install printer and associate port
    Add-PrinterDriver  -ComputerName $ComputerName -Name $DRIVERNAME -ErrorAction "silentlycontinue"
    Add-PrinterPort -ComputerName $ComputerName -Name $PRINTERIP -PrinterHostAddress $PRINTERIP -ErrorAction "silentlycontinue"
    Add-Printer -ComputerName $ComputerName -Name $PRINTERNAME -DriverName $DRIVERNAME -PortName $PRINTERIP -ErrorAction "silentlycontinue"
}
