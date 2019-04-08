@ECHO OFF 
:: This batch file reveals OS, hardware settings through networking.
TITLE My System Info
ECHO Please wait... Checking nodes system information.
:: Section 1: OS information.
ECHO  ===================================================
ECHO       Script developed by Celiz Matias 100286315 
ECHO  ===================================================
wmic /node:@BuenosAiresNodes.txt /output:SystemInformation.csv ComputerSystem get TotalPhysicalMemory,model,systemtype,usernameclear /format:csv
wmic /node:@BuenosAiresNodes.txt /output:HardwareInformation.csv bios get serialNumber /format:csv
 
ECHO  ===================================================
ECHO   The information is now available in out.csv file!
ECHO  ===================================================