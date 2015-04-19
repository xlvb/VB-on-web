@echo off 
echo Checking your system infor, Please wating... 
systeminfo | findstr /c:"Host Name" 
systeminfo | findstr /c:"Domain" 
systeminfo | findstr /c:"OS Name" 
systeminfo | findstr /c:"OS Version" 
systeminfo | findstr /c:"System Manufacturer" 
systeminfo | findstr /c:"System Model" 
systeminfo | findstr /c:"System type" 
systeminfo | findstr /c:"Total Physical Memory" 
ipconfig | findstr IPv4
ipconfig | findstr IPv6

echo.

echo Hard Drive Space: 
wmic diskdrive get size

echo. 
echo.

echo Service Tag: 
wmic bios get serialnumber

echo. 
echo. 
echo CPU: 
wmic cpu get name

echo Completed! Thank you!

pause
