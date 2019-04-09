@ECHO OFF 
:: This batch file reveals OS, hardware settings through networking.
TITLE My System Info
ECHO Please wait... Checking nodes system information.
ECHO  ===================================================
ECHO       Script developed by Celiz Matias 100286315 
ECHO  ===================================================
:: Total physical memory installed, node model, system type and last username logged
wmic /node:@NeuquenNodes.txt /output:SystemInformation.csv ComputerSystem get Name,TotalPhysicalMemory,Model,SystemType,UserName /format:csv
:: Serial Number information
wmic /node:@NeuquenNodes.txt /output:HardwareInformation.csv bios get serialNumber /format:csv
:: Memory information
wmic /node:@NeuquenNodes.txt /output:MemoryInformation.csv MemoryChip get Manufacturer,Capacity,PartNumber,Speed,MemoryType,DeviceLocator,FormFactor /format:csv
ECHO  ===================================================
ECHO   The information is now available in out.csv file!
ECHO  ===================================================