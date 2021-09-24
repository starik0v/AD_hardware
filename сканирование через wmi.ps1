$OS = gwmi  Win32_OperatingSystem | Select Caption, OSArchitecture, CSName #TotalVisibleMemorySize 
$CPU = gwmi  Win32_Processor | Select Name #Architecture, DeviceID
$RAM = gwmi  Win32_MemoryDevice | Select DeviceID, StartingAddress, EndingAddress
$MB = gwmi  Win32_BaseBoard | Select Manufacturer, Product, Version
$VGA = gwmi  Win32_VideoController | Select Name, AdapterRam
$HDD = gwmi  Win32_DiskDrive | select Model, Size
$Volumes = gwmi  Win32_LogicalDisk -Filter "MediaType = 12" | Select DeviceID, Size, FreeSpace
$IP = gwmi Win32_NetworkAdapterConfiguration | Select -ExpandProperty IPAddress #-Filter IPEnabled=$true
$MAC = gwmi Win32_NetworkAdapter |? {$_.NetConnectionID -ne $null} | Select MACaddress
$Dom = gwmi Win32_ComputerSystem | Select Domain
$Mon = gwmi Win32_DesktopMonitor |Select Name, MonitorManufacturer
#$Mon = get-systeminfo -properties monitormanuf, monitorname
#$OfVer = gwmi Win32_Product |select Version
#$OfVer = reg query "HKEY_CLASSES_ROOT\Word.Application\CurVer"
$OfVer = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer").'(default)'
$ie = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Internet Explorer').Version

#"Computer Name: `n`t" + $OS.CSName + "`n"
#"Operating System: `n`t" + $OS.Caption + " " + $OS.OSArchitecture + "`n" 


$(
"МАС адрес:            " + $MAC.MACaddress + "`n"
"HDD:                  " + $HDD.Model + " " + $HDD.Size + "`n" 
"IP-address:           " + $IP + "`n"
"Рабочая группа/домен: " + $Dom.Domain + "`n"
"Процессор:            " + $CPU.Name + "`n"
"Материнская плата:    " + $MB.Manufacturer + " " + $MB.Product + " " + $MB.Version + "`n"
"Системная память:     " + $RAM.DeviceId + " " + $RAM.StartingAddress + " " + $RAM.EndingAddress + "`n"
"Видеоадаптер:         " + $VGA.Name + " " + $VGA.AdapterRam + "`n"
"Жесткий диск:         " + $HDD.Model + " " + $HDD.Size + "`n"
"Операционная система: " + $OS.Caption + " " + $OS.OSArchitecture + "`n"
"Internet Explorer:    " + $ie + "`n"
"MS Office:            " + $OfVer + "`n"
"
 Office 97   -  7.0
 Office 98   -  8.0
 Office 2000 -  9.0
 Office XP   - 10.0
 Office 2003 - 11.0
 Office 2007 - 12.0
 Office 2010 - 14.0 
 Office 2013 - 15.0
 Office 2016+ - 16.0"
) *>&1 > "C:\Users\istarikov\Desktop\systeminfo4.txt"

#Get-Content -Path C:\Users\istarikov\Desktop\systeminfo3.txt  
#"$IE, $Dom, $Mon, $OfVer, $OS, $CPU, $RAM, $MB, $VGA, $HDD, $Volumes, $IP, $MAC"| out-file -filepath C:\Users\istarikov\Desktop\systeminfo3.txt #-Width 200 
#"$MAC, $HDD, $IP, $Dom, $CPU, $MB, $RAM, $VGA, $Mon, $HDD, $OS" | out-file -filepath C:\Users\istarikov\Desktop\systeminfo3.txt -width 200 