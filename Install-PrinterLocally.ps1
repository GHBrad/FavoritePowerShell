   
#Created by Brad Johnson
#Created 10/4/2017

[CmdletBinding()]
    Param
    (
        
        [Parameter(Mandatory=$true)]
        [String]
        $ComputerName,
        [Parameter(Mandatory=$false)]
        [String]
        $PRINTERNAME="<PRINTER NAME>",
        [Parameter(Mandatory=$false)]
        [String]
        $PRINTERIP="<PRINTER IP ADDRESS>",
        [Parameter(Mandatory=$false)]
        [String]
        $DRIVERNAME="<PRINTER DRIVER NAME FOUND IN THE INF FILE>",
        [Parameter(Mandatory=$false)]
        [String]
        $DRIVERPATH="<FOLDER PATH OF DRIVER>"
    )

#If the name is blank, use localhost
$InstallingLocally = $FALSE
if ($ComputerName -eq ""){
    $ComputerName = $env:COMPUTERNAME
    $InstallingLocally = $TRUE
}

# Remove printer if already installed
Remove-Printer -ComputerName $ComputerName -Name $PRINTERNAME
Remove-PrinterPort -computername $computername -Name $PRINTERIP

# Copy and install drivers to machine
robocopy "\\nasty\share\downloads\Drivers\Windows\Printers\$DRIVERPATH" "\\$ComputerName\c$\temp\PRINTER DRIVERS" /mir
if ($InstallingLocally){ #Installing locally. Run pnputil.exe directly.
   pnputil.exe /a "C:\temp\PRINTER DRIVERS\OEMSETUP.INF"
}else{ #Installing remotely. Run pnputil.exe remotely via Invoke-Command.
   Invoke-Command -ComputerName $ComputerName -ScriptBlock {pnputil.exe /a "C:\temp\PRINTER DRIVERS\OEMSETUP.INF"}
}

# Install printer and associate port
Add-PrinterDriver  -ComputerName $ComputerName -Name $DRIVERNAME
Add-PrinterPort -ComputerName $ComputerName -Name $PRINTERIP -PrinterHostAddress $PRINTERIP
Add-Printer -ComputerName $ComputerName -Name $PRINTERNAME -DriverName $DRIVERNAME -PortName $PRINTERIP

