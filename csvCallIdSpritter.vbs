Option Explicit

Dim fso, ts, line, headers, data, folderPath, newFilePath
Set fso = CreateObject("Scripting.FileSystemObject")

Dim file

Dim scriptPath : scriptPath = WScript.ScriptFullName
Dim scriptDir : scriptDir = fso.GetParentFolderName(scriptPath)

For Each file In fso.GetFolder(scriptDir).Files
    If LCase(fso.GetExtensionName(file)) = "csv" Then
        ' ファイル開く
        Set ts = fso.OpenTextFile(file.Path)

        ' Header 読み込み
        line = ts.ReadLine
        headers = Split(line, ",")

        ' メモリにデータ格納
        Set data = CreateObject("Scripting.Dictionary")

        ' 複数ファイルを読み込む
        Do Until ts.AtEndOfStream
            line = ts.ReadLine
            Dim items : items = Split(line, ",")

            ' フォルダとファイル名のフォーマット
            Dim firstColumnValue : firstColumnValue = items(0)
            Dim dateTimeValue : dateTimeValue = Split(items(1), " ")

            ' file name
            Dim fileDateValue : fileDateValue = Replace(dateTimeValue(0), "-", "")

            ' folder name
            Dim dateParts : dateParts = Split(dateTimeValue(0), "-")
            Dim folderDateValue : folderDateValue = dateParts(1) & dateParts(2)

            folderPath = scriptDir & "\" & folderDateValue
            newFilePath = folderPath & "\" & fileDateValue & "_" & firstColumnValue & ".csv"

            ' フォルダの存在確認
            If Not fso.FolderExists(folderPath) Then
                fso.CreateFolder(folderPath)
            End If

            ' 追加
            If Not data.Exists(newFilePath) Then
                data.Add newFilePath, Join(headers, ",") & vbCrLf & line
            Else
                data(newFilePath) = data(newFilePath) & vbCrLf & line
            End If
        Loop

        ' 閉じる
        ts.Close
    End If
Next

' 全てのファイルを閉じる
Dim key
For Each key In data.Keys
    Dim newFile : Set newFile = fso.CreateTextFile(key, True)
    newFile.Write data(key)
    newFile.Close
Next

MsgBox "Completed."