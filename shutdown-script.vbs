' Declare variables and create a wscript object that runs command-prompt commands.

Dim X,S
Dim varcheck
Dim objWshShl 
Set objWshShl = WScript.CreateObject("wscript.shell")

' Warn the user when the program wakes up from 3600000 milliseconds (1 hour) of sleep, and if the user clicks "yes" button, shut down the PC else, warn the user 1 hour later to shut down the PC.

Do
varcheck = 0
WScript.Sleep 3600000
X = MsgBox("Hey, you used this computer for more than 1 hour. Please stop your computer addiction and immediately shut down this computer.", 4+48+4096, "shutdown-script")
If X = 6 Then
        S=MsgBox("Confirm shutdown", 4+48+4096, "shutdown-script")
Else
        varcheck = 1
End If

If S = 6 Then 
	objWshShl.Run "Shutdown /s /t 1",0
Else
varcheck = 1
End If

Loop While varcheck = 1