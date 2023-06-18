Option Explicit

Dim fso, ts, line, headers, data, folderPath, newFilePath
Set fso = CreateObject("Scripting.FileSystemObject")

Dim file

' Get the script's own location.
Dim scriptPath : scriptPath = WScript.ScriptFullName
Dim scriptDir : scriptDir = fso.GetParentFolderName(scriptPath)

' Create a Dictionary to hold the data.
Set data = CreateObject("Scripting.Dictionary")

' Read each file in the same directory as the script.
For Each file In fso.GetFolder(scriptDir).Files
    If LCase(fso.GetExtensionName(file)) = "csv" Then
        ' Open the file.
        Set ts = fso.OpenTextFile(file.Path)
    
        ' Read the headers.
        line = ts.ReadLine
        headers = Split(line, ",")
    
        ' Read the rest of the file.
        Do Until ts.AtEndOfStream
            line = ts.ReadLine
            Dim items : items = Split(line, ",")

            ' Format the file name and folder path.
            Dim firstColumnValue : firstColumnValue = items(0)
            Dim dateTimeValue : dateTimeValue = Split(items(1), " ") ' Split date and time.
            Dim dateValue : dateValue = Replace(dateTimeValue(0), "-", "")
            folderPath = scriptDir & "\" & dateValue
            newFilePath = folderPath & "\" & dateValue & "_" & firstColumnValue & ".csv"
    
            ' Create the folder if it doesn't exist.
            If Not fso.FolderExists(folderPath) Then
                fso.CreateFolder(folderPath)
            End If
    
            ' Add the line to the appropriate file.
            If Not data.Exists(newFilePath) Then
                data.Add newFilePath, Join(headers, ",") & vbCrLf & line
            Else
                data(newFilePath) = data(newFilePath) & vbCrLf & line
            End If
        Loop
    
        ' Clean up.
        ts.Close
    End If
Next

' Write all data to the files.
Dim key
For Each key In data.Keys
    Dim newFile : Set newFile = fso.CreateTextFile(key, True)
    newFile.Write data(key)
    newFile.Close
Next

' Display a completion message.
MsgBox "Completed."
