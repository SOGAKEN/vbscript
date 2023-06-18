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
            Dim dateTimeValue : dateTimeValue = Split(items(1), " ") ' Split date and time.
            Dim dateValue : dateValue = Replace(dateTimeValue(0), "-", "")
            folderPath = scriptDir & "\" & dateValue
            newFilePath = folderPath & "\" & dateValue & "_" & firstColumnValue & ".csv"
    
            ' Create the folder if it doesn't exist.
            If Not fso.FolderExists(folderPath) Then
                fso.CreateFolder(folderPath)
            End If

            ' Add the line to the appropriate file.
            Dim newFile
            Dim counter : counter = 1
            Dim originalFilePath : originalFilePath = newFilePath
            While fso.FileExists(newFilePath) ' Check if file already exists
                newFilePath = Left(originalFilePath, InStrRev(originalFilePath, ".") - 1) & "_" & counter & ".csv"
                counter = counter + 1
            Wend
            If Not data.Exists(newFilePath) Then
                Set newFile = fso.CreateTextFile(newFilePath, True)
                data.Add newFilePath, newFile
                newFile.WriteLine(Join(headers, ","))
            Else
                Set newFile = data(newFilePath)
            End If
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
