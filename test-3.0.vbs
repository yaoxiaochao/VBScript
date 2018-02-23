Sub Main()




Dim objFile, File
Dim objClip
Dim username, password, filename, scriptname, Script



Set objFile = CreateObject("Scripting.FileSystemObject")
filename = xsh.Dialog.Prompt ("Filename with full path", "Filename", "", 0)
Set File = objFile.OpenTextFile(filename, 1)

Set objFile = CreateObject("Scripting.FileSystemObject")
scriptname = xsh.Dialog.Prompt ("Filename with full path", "Scriptname", "", 0)
Set Script = objFile.OpenTextFile(scriptname, 1)


username = xsh.Dialog.Prompt ("Username", "User name", "", 0)
password = xsh.Dialog.Prompt ("Password", "Password", "", 1)



dim devicename,row,a

Do Until File.AtEndOfStream
hostname = File.Readline
'xsh.Screen.Send(hostname + VbCr)
xsh.Screen.Send("ssh "+username+"@" + hostname + VbCr)

xsh.Screen.WaitForString "(yes/no)?"
xsh.Screen.Send "yes" & VbCr
xsh.Session.Sleep 1000
xsh.Screen.Send(password + VbCr)
xsh.Session.Sleep 1000
xsh.Screen.Send "N" & VbCr
xsh.Session.Sleep 1000

row = xsh.Screen.CurrentRow
devicename = xsh.Screen.Get(row,1,row,30)
a = len(devicename)
b = a - 1
c = 2
d = xsh.Screen.Get(row,1,row,1)
if d = "H" Then 
	c = 7
end if
devicename = xsh.Screen.Get(row,c,row,b)
xsh.Session.LogFilePath = "E:\work\log\"+devicename+"_%Y-%m-%d_%t.log"
xsh.Session.Sleep 1000
xsh.Session.StartLog

if hotstname > 0 Then Exit do

Set objFile = CreateObject("Scripting.FileSystemObject")
Set Script = objFile.OpenTextFile(scriptname, 1)

Do Until Script.AtEndOfStream
command = Script.Readline
'xsh.Screen.Send(command + VbCr)
xsh.Screen.Send(command + VbCr)
xsh.Session.Sleep 3000
Loop

Script.Close
xsh.Session.Sleep 1000
xsh.Session.StopLog

Loop 


File.Close





End Sub


