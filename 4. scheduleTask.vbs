Dim WshShell, mainDir, Filename, args
Set WshShell = CreateObject("WScript.Shell")
mainDir    = WshShell.CurrentDirectory

Filename = """" & mainDir & "\3. countDownBG.vbs" & """"

args = "-noexit -file " & mainDir & "\scheduleTask.ps1" & " " & Filename

Set app = CreateObject("Shell.Application")
app.ShellExecute "powershell.exe", args, , "runas", 0
