Sub Main()


xsh.Session.Open "E:\work\Ñ²¼ì\Ñ²¼ì½Å±¾\U2000.xsh"


Dim objFile, File
Dim objClip
Dim username, password, filename

Set objFile = CreateObject("Scripting.FileSystemObject")
filename = xsh.Dialog.Prompt ("Filename with full path", "Filename", "", 0)
'Set File = objFile.OpenTextFile("d:\test.txt", 1)
Set File = objFile.OpenTextFile(filename, 1)

username = xsh.Dialog.Prompt ("Username", "User name", "", 0)
password = xsh.Dialog.Prompt ("Password", "Password", "", 1)

xsh.Session.LogFilePath("E:\work\log\test_%Y-%m-%d_%t.log")

xsh.Session.StartLog



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
xsh.Screen.Send "quit" & VbCr
xsh.Session.Sleep 1000
Loop

File.Close

xsh.Session.StopLog

End Sub


