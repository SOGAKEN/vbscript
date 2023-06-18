Option Explicit

Dim fso, ts, line, headers, data, folderPath, newFilePath, counter
Set fso = CreateObject("Scripting.FileSystemObject")

Dim file

' Get the script's own location.
Dim scriptPath : scriptPath = WScript.ScriptFullName
Dim scriptDir : scriptDir = fso.GetParentFolderName(scriptPath)

' Read each file in the same directory as the script.
For Each file In fso.GetFolder(scriptDir).Files
    If LCase(fso.GetExtensionName(file)) = "csv" Then
        ' Open the file.
        Set ts = fso.OpenTextFile(file.Path)
    
        ' Read the headers.
        line = ts.ReadLine
        headers = Split(line, ",")
    
        ' Create a Dictionary to hold the data.
        Set data = CreateObject("Scripting.Dictionary")
    
        ' Read the rest of the file.
        Do Until ts.AtEndOfStream
            line = ts.ReadLine
            Dim items : items = Split(line, ",")
    
            ' Format the file name and folder path.
            Dim firstColumnValue : firstColumnValue = items(0)
            Dim dateTimeValue : dateTimeValue = Split(items(1), " ") ' Split date and time.
            Dim dateValue : dateValue = Replace(dateTimeValue(0), "-", "")
            folderPath = scriptDir & "\" & dateValue
            newFilePath = folderPath & "\" & dateValue & "_" & firstColumnValue
            Dim suffix : suffix = ""
            counter = 1
            While fso.FileExists(newFilePath & suffix & ".csv")
                counter = counter + 1
                suffix = "_" & counter
            Wend
            newFilePath = newFilePath & suffix & ".csv"
    
            ' Create the folder if it doesn't exist.
            If Not fso.FolderExists(folderPath) Then
                fso.CreateFolder(folderPath)
            End If
    
            ' Check if the value already exists in the dictionary.
            If Not data.Exists(firstColumnValue) Then
                ' Add the value to the dictionary.
                data.Add firstColumnValue, newFilePath
            Else
                ' The value already exists, so append a number to the file name.
                newFilePath = Left(data.Item(firstColumnValue), Len(data.Item(firstColumnValue)) - 5) & "_" & counter & ".csv"
            End If
    
            ' Add the line to the appropriate file.
            Dim newFile
            Set newFile = fso.CreateTextFile(newFilePath, True)
            newFile.WriteLine(Join(headers, ","))
            newFile.WriteLine line
            newFile.Close
        Loop
    
        ' Clean up.
        ts.Close
    End If
Next

' Display a completion message.
MsgBox "Completed."
