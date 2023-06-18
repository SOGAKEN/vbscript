Option Explicit

Dim fso, line, headers, data, folderPath, newFilePath
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the script's own location.
Dim scriptPath : scriptPath = WScript.ScriptFullName
Dim scriptDir : scriptDir = fso.GetParentFolderName(scriptPath)

' Start processing files.
ProcessFiles scriptDir

' Display a completion message.
MsgBox "Completed."

Sub ProcessFiles(folderPath)
    Dim file, folder

    ' Read each file in the current directory.
    For Each file In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(file)) = "csv" Then
            ' Process the file.
            ProcessFile file.Path
        End If
    Next

    ' Recursively process each subdirectory.
    For Each folder In fso.GetFolder(folderPath).SubFolders
        ProcessFiles folder.Path
    Next
End Sub

Sub ProcessFile(filePath)
    ' Open the file.
    Dim ts : Set ts = fso.OpenTextFile(filePath)
    
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
            
        ' For file name
        Dim fileDateValue : fileDateValue = Replace(dateTimeValue(0), "-", "")

        ' For folder name
        Dim dateParts : dateParts = Split(dateTimeValue(0), "-")
        Dim folderDateValue : folderDateValue = dateParts(1) & dateParts(2) ' or use dateParts(2) for "19"

        folderPath = scriptDir & "\" & folderDateValue
        newFilePath = folderPath & "\" & fileDateValue & "_" & firstColumnValue & ".csv"
    
        ' Create the folder if it doesn't exist.
        If Not fso.FolderExists(folderPath) Then
            fso.CreateFolder(folderPath)
        End If
    
        ' Add the line to the appropriate file.
        Dim newFile
        If Not data.Exists(newFilePath) Then
            Set newFile = fso.CreateTextFile(newFilePath, True)
            data.Add(newFilePath, newFile)
            newFile.WriteLine Join(headers, ",")
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
End Sub
