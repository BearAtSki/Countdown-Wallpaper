Option Explicit

Const ForAppending = 8
Const TristateFalse = 0
Const Overwrite = True

Const WindowsFolder = 0
Const SystemFolder = 1
Const TemporaryFolder = 2

Dim FileSystem
Dim Filename, OldText, NewText
Dim OriginalFile, TempFile, Line
Dim TempFilename

Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir    = WshShell.CurrentDirectory


' If, for whatever reason you change the name of countDownBG.vbs, change Filename to that.

Filename = "3. countDownBG.vbs"


OldText = "mainDir ="
NewText = "mainDir = """ & strCurDir & """"

Set FileSystem = CreateObject("Scripting.FileSystemObject")
Dim tempFolder: tempFolder = FileSystem.GetSpecialFolder(TemporaryFolder)
TempFilename = FileSystem.GetTempName

If FileSystem.FileExists(TempFilename) Then
    FileSystem.DeleteFile TempFilename
End If

Set TempFile = FileSystem.CreateTextFile(TempFilename, Overwrite, TristateFalse)
Set OriginalFile = FileSystem.OpenTextFile(Filename)

Do Until OriginalFile.AtEndOfStream
    Line = OriginalFile.ReadLine

    If InStr(Line, OldText) > 0 Then
        Line = NewText
    End If 

    TempFile.WriteLine(Line)
Loop

OriginalFile.Close
TempFile.Close

FileSystem.DeleteFile Filename
FileSystem.MoveFile TempFilename, Filename

Wscript.Quit