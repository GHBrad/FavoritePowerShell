#Created by Brad Johnson
#Created 2/19/2019

#Purpose: Install all printers
#Source file for printers to install: \\<NETWORK SHARE\Printers.csv

#Capture destination computer
$ComputerName = Read-Host "Enter computer name:"

#Import CSV file
$Printers = Import-Csv "\\<NETWORK SHARE\Printers.csv"

#Loop through each row containing printer details in the CSV file
foreach ($Printer in $Printers) {
    #Read printer data from each field in each row and assign the data to a variable as below
    
    $PrinterName = $Printer.Name
    $PrinterIP = $Printer.IP
    $DriverName = $Printer.DriverName
    $DriverPath = $Printer.DriverPath


    # Remove printer if already installed
    Remove-Printer -ComputerName $ComputerName -Name $PRINTERNAME
    Remove-PrinterPort -computername $computername -Name $PRINTERIP

    # Copy and install drivers to machine
    robocopy "\\<NETWORK SHARE>\$DRIVERPATH" "\\$ComputerName\c$\temp\PRINTER DRIVERS" /mir
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {pnputil.exe /a "C:\temp\PRINTER DRIVERS\OEMSETUP.INF"}

    # Install printer and associate port
    Add-PrinterDriver  -ComputerName $ComputerName -Name $DRIVERNAME
    Add-PrinterPort -ComputerName $ComputerName -Name $PRINTERIP -PrinterHostAddress $PRINTERIP
    Add-Printer -ComputerName $ComputerName -Name $PRINTERNAME -DriverName $DRIVERNAME -PortName $PRINTERIP
}
