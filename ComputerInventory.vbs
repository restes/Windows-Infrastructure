On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

'change this value to the IP address or hostname of the machine you need to audit
strIPvalue = "localhost"

CALL GenerateReport(strIPvalue)

WScript.Echo "Inventory Complete "

'=================================================================================
'SUB-ROUTINE GenerateReport
SUB GenerateReport(strIPvalue)

'Script to change a filename using timestamps
strPath = "\\<path of your>\<output>\" 'Change the path to appropriate value
strMonth = DatePart("m", Now())
strDay = DatePart("d",Now())

if Len(strMonth)=1 then
strMonth = "0" & strMonth
else
strMonth = strMonth
end if

if Len(strDay)=1 then
strDay = "0" & strDay
else
strDay = strDay
end if

strFileName = DatePart("yyyy",Now()) & strMonth & strDay
strFileName = Replace(strFileName,":","")
'=================================================================================

'Variable Declarations
Const ForAppending = 8

'===============================================================================
'Main Body
On Error Resume Next

'Get COMPUTER NAME Information
strComputer = strIPvalue
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
For Each objItem in colItems
CompName = objItem.SystemName
Next

'Get ASSET ID Information
strAssetID = InputBox("Please enter the Asset ID tag #:", _
    "Need Asset ID #")

'Get USER Information
strUserID = InputBox("Please enter the name of the person assigned to this computer:", _
    "Enter the first and last name of the user.")

'See if the file name exists; quit running if it does.
Set objFSO = CreateObject("Scripting.FileSystemObject")
if objFSO.FileExists(strPath & strAssetID & "_" & CompName & "_" & strUserID & ".txt") then
WScript.Quit
end if

'Set the file location to collect the data.
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(strPath & strAssetID & "_" & CompName & "_" & strUserID & ".txt")
Wscript.Echo("File created successfully.")

Set objTextFile = nothing
Set objTextFile = objFSO.OpenTextFile(strPath & strAssetID & "_" & CompName & "_" & strUserID & ".txt", ForAppending, True)
WScript.Echo("File opened successfully.")


''==============================================================
'Print HEADER
objTextFile.Write "================================================================" & VBCRLF & VBCRLF
objTextFile.Write " WORKSTATION/LAPTOP RESOURCE INVENTORY REPORT " & VBCRLF
objTextFile.Write " DATE: " & FormatDateTime(Now(),1) & " " & VBCRLF
objTextFile.Write " TIME: " & FormatDateTime(Now(),3) & " " & VBCRLF & VBCRLF
objTextFile.Write "================================================================" & VBCRLF & VBCRLF


objTextFile.Write "COMPUTER INFORMATION" & VBCRLF
objTextFile.Write "================================================================" & VBCRLF
'==============================================================

'Get the Asset ID to file.
objTextFile.Write "ASSET ID: " & strAssetID & VBCRLF
WScript.Echo("Asset ID logged successfully.")

'Write the User ID to file.
objTextFile.Write "NAME OF USER: " & strUserID & VBCRLF
WScript.Echo("User ID logged successfully.")

'Write COMPUTER NAME to file.
objTextFile.Write "COMPUTER NAME: " & CompName & VBCRLF
WScript.Echo("Computer name logged successfully.")

'Get MAKE & MODEL Information
Set colItems = objWMIService.ExecQuery("Select * from Win32_computersystem")
For Each objItem in colItems
objTextFile.Write "MAKE: " & objItem.Manufacturer & VBCRLF
objTextFile.Write "MODEL: " & objItem.Model & VBCRLF
WScript.Echo("Make and model logged successfully.")
Next

'Get OS Information
Set colSettings = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objOperatingSystem in colSettings
objTextFile.Write "OS: " & objOperatingSystem.Name & VBCRLF
objTextFile.Write "ARCHITECTURE: " & objOperatingSystem.OSArchitecture & VBCRLF
objTextFile.Write "SERVICE PACK VERSION: " & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion & VBCRLF
WScript.Echo("OS, Architecture, Service Pack logged successfully.")
Next

'Get Total Physical memory
Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings
objTextFile.Write "TOTAL PHYSICAL RAM: " & Round((objComputer.TotalPhysicalMemory/1000000000),4) & " GB" & VBCRLF
WScript.Echo("Total RAM logged successfully.")
Next

'Get MS Office version
objTextFile.Write " " & VBCRLF & "MICROSOFT OFFICE VERSIONS:" & VBCRLF
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colApps = objWMIService.ExecQuery _
   ("Select * from Win32_Product Where Caption Like '%Microsoft Office%'")
For Each objApp in colApps
   objTextFile.Write " " & objApp.Caption & " " & objApp.Version & VBCRLF
Next

WScript.Echo("MS Office versions logged successfully.")

'Get Adobe version
objTextFile.Write " " & VBCRLF & "ADOBE VERSIONS:" & VBCRLF
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colApps = objWMIService.ExecQuery _
   ("Select * from Win32_Product Where Caption Like '%Adobe%'")
For Each objApp in colApps
   objTextFile.Write " " & objApp.Caption & " " & objApp.Version & VBCRLF
Next

WScript.Echo("Adobe versions logged successfully.")

'===========================================
'Close text file after writing logs
objTextFile.Write VbCrLf
objTextFile.Close

'===========================================

'Clean Up
SET colIESettings=NOTHING
SET colItems=NOTHING
SET colSettings=NOTHING
SET colDisks=NOTHING
SET AdapterSet=NOTHING
SET objWMIService=NOTHING
SET objWMIService2=NOTHING
SET objFSO=NOTHING
SET objTextFile=NOTHING

END SUB
'===================================================================
