Dim WshShell,FSO,secs

Set WshShell = WScript.CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

secs=180
currentdir=Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
kill_file=currentdir+"New Text Document.txt"

Do While true
	if FSO.FileExists(kill_file) then exit do
	WScript.Sleep(secs*1000) 'secs in miliseconds so multiply with 1000
	WshShell.SendKeys "{SCROLLLOCK}"
	WshShell.SendKeys "{SCROLLLOCK}"
Loop
