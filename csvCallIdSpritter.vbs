Option Explicit

Dim fso, ts, line, headers, data, folderPath, newFilePath
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
            folderPath = scriptDir & "\" & firstColumnValue
            Dim counter : counter = 1
            newFilePath = folderPath & "\" & firstColumnValue & ".csv"
            
            ' If the file already exists, add a counter to the filename.
            While fso.FileExists(newFilePath)
                newFilePath = folderPath & "\" & firstColumnValue & "_" & counter & ".csv"
                counter = counter + 1
            Wend
            
            ' Create the folder if it doesn't exist.
            If Not fso.FolderExists(folderPath) Then
                fso.CreateFolder(folderPath)
            End If
            
            ' Add the line to the appropriate file.
            Dim newFile
            Set newFile = fso.CreateTextFile(newFilePath, True)
            data.Add newFilePath, newFile
            newFile.WriteLine(Join(headers, ","))
            newFile.WriteLine line
        Loop
        
        ' Clean up.
        ts.Close
        Dim key
        For Each key In data
            data(key).Close
        Next
    End If
Next
