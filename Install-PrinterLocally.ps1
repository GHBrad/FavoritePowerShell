   
#Created by Brad Johnson
#Created 10/4/2017
#I haven't been able to figure out how to set a relative path to install the INF file using pnputil, Add-WindowsDriver or Add-PrinterDriver. So,
#if you're going to use this script as a template for other printer installs, change the printer named folder ('c:\temp\<PRINTER NAME>\OEMSETUP.INF) to the appropriate folder name.


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


        # Remove printer if already installed
        Remove-Printer -ComputerName $ComputerName -Name $PRINTERNAME
        Remove-PrinterPort -computername $computername -Name $PRINTERIP

        # Copy and install drivers to machine
        robocopy "\\<NETWORK SHARE>\$DRIVERPATH" "\\$ComputerName\c$\temp\$PRINTERNAME" /mir
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {pnputil.exe /a 'C:\temp\<PRINTER NAME>\OEMSETUP.INF'}

        # Install printer and associate port
        Add-PrinterDriver  -ComputerName $ComputerName -Name $DRIVERNAME
        Add-PrinterPort -ComputerName $ComputerName -Name $PRINTERIP -PrinterHostAddress $PRINTERIP
        Add-Printer -ComputerName $ComputerName -Name $PRINTERNAME -DriverName $DRIVERNAME -PortName $PRINTERIP

