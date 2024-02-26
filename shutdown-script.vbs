' Declare global variables and objects that can be used through functions
Dim scriptDir, file, config
Set objWshShl = CreateObject("wscript.shell")
Set fso = CreateObject("Scripting.FileSystemObject") 
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
Set file = fso.OpenTextFile(scriptDir & "\config.txt", 1)

' Create a class to encapsulate every line in config.txt to different variables inside class
Class configurations
        Public milliseconds
	Public autostart
End Class
Set config = new configurations

main
Sub main()
	' Store the number from milliseconds.txt in the variable "file_content" and use it as the argument for the object Wscript.Sleep
	setconfig()

	' If autostart enabled, create a registry key to autostart script everytime user turn on the PC; else if autostart disabled, remove the registry key for autostarting the script.
	If config.autostart Then 
		startup()
	Else 
                objWshShl.Run "reg delete ""HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"" /v ""shutdown-scheduler"" /f",0
	End If

        ' Do a simple calculation to convert milliseconds into hours and minutes, and use the one that is easier to read.
        dim result, time, hours, minutes, seconds, i
        time = 0 : i = 1
        incrementresult time, result, hours, minutes, seconds, i

        ' Warn the user when the program wakes up after 3600000 (can be changed through the milliseconds.txt or set-new-timer.bat file) milliseconds (1 hour) of sleep, and if the user clicks the "yes" button, shutdown the PC; otherwise,         warn the user 1 hour later to shut down the PC.
	Dim dialoguebox1, dialoguebox2
        Do While True
                WScript.Sleep config.milliseconds
                dialoguebox1 = MsgBox ("Hey, you used this computer for " & time & " " & result & ". Please stop your computer addiction and immediately shut down this computer.", 4+48+4096, "shutdown-script")
                incrementresult time, result, hours, minutes, seconds, i

                If dialoguebox1 = 6 Then
                        dialoguebox2 = MsgBox("Confirm shutdown", 4+48+4096, "shutdown-script")
                End If

                If dialoguebox2 = 6 Then
			set fso = nothing
	                objWshShl.Run "Shutdown /s /t 0",0
			set objWshShl = nothing
                End If
        Loop
End Sub

Sub setconfig()
	' Source https://stackoverflow.com/questions/29320616/how-does-one-declare-an-array-in-vbscript
	Dim file_contents(2)

	For i = 1 To 2
	        file_contents(i) = file.readline
	Next

	config.milliseconds = file_contents(1)
	config.autostart = file_contents(2)

	file.close
	Set file = nothing
End Sub

Sub startup()
        Const HKEY_CURRENT_USER = &H80000001

        strComputer = "."
 
        Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
 
        strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        strValueName = "shutdown-scheduler"
        strValue = scriptDir & "\shutdown-script.vbs"
        oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
        Set oReg = nothing
End Sub

Sub incrementresult(time, result, hours, minutes, seconds, i)
        hours = (config.milliseconds * i) / 3600000
        minutes = (config.milliseconds * i) / 60000
        seconds = (config.milliseconds * i) / 1000

        If hours => 1 Then
                result = "hours"
                time = hours
        End If
        If minutes => 1 And 60 > minutes Then
                result = "minute"
                time = minutes
        End If
        If seconds < 60 Then
                result = "second"
                time = seconds
        End If

        i = i + 1
End Sub
