Dim fso, folder, file, folderName, fileName, fileExt, newFileName, FilesInFolder, counter, sc
Dim dateMonth, dateDay, dateYear, dateStart, dateNew
Dim WshShell

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")  


' ---------------------------------------------------------------------------
' ------------- CHANGE THE START DATE BEFORE RUNNING THE SCRIPT -------------
' ---------------------------------------------------------------------------

' Start date for the countdown - "jan-31-2023" and "01-january-2023" works just fine
dateStart="feb-28-2023"

' Desired File Name for the Countdown images
fileName = "countDown"

' Folder containing the CountDown images
folderName = "cd"

' File extension - should be PNG out of After Effects.
fileExt = "png"

' ---------------------------------------------------------------------------
' ---------------------------------------------------------------------------
' ---------------------------------------------------------------------------


folderName = WshShell.CurrentDirectory & "\" & folderName

Set folder = fso.GetFolder(folderName)

counter = 0

sc = 1
dateNew = dateStart

' Counting all files in folder with defined extension
For Each File In folder.Files
   if LCase(fso.GetExtensionName(file.Name)) = fileExt Then
     counter = counter + 1
   End If  
Next

' Loop over all files in the folder
For Each file In folder.Files
	' Check for files with the specified extension
	If sC <= counter Then
		If LCase(fso.GetExtensionName(file.Name)) = fileExt Then
			' Converting single digit month and day to two digits
			dateMonth = Right(String(2, "0") & Month(dateNew), 2)
			dateDay = Right(String(2, "0") & Day(dateNew), 2)
			
			dateYear = Right(Year(dateNew), 2)
			
			' File name from Prefix + _ + Year + Month + Day + Extension
			newFileName = fileName & "_" & dateYear & dateMonth & dateDay & "." & fileExt
			
			file.Name = newFileName
				
			dateNew=DateAdd("d",1,dateNew)
			sC = sC+1
		End If
	Else

		Exit For
		
	End If	
	
Next