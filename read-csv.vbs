

Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

Set inputFile = oFSO.OpenTextFile(".\test.csv")

Dim isFirstLine
isFirstLine = True
Do While inputFile.AtEndOfStream <> True
    arr = Split(inputFile.ReadLine, ",")

    If isFirstLine Then
        MsgBox("header " & arr(0))
        
        isFirstLine = False
    Else
        MsgBox(arr(0))
    End If
Loop

Set inputFile = Nothing