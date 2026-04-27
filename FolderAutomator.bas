Attribute VB_Name = "FolderAutomator"
' ====================================================================
' Tool: Document Folder Automator
' Purpose: Creates folder structures from Excel lists
' ====================================================================
Sub CreateFolderStructure()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim rootPath As String, folderName As String
    Dim createdCount As Integer

    Set ws = ThisWorkbook.Sheets("Folder_Automator")
    rootPath = ws.Range("B2").Value ' Cell where user puts the main path
    
    If Right(rootPath, 1) <> "\" Then rootPath = rootPath & "\"
    
    ' Check if root path exists
    If Dir(rootPath, vbDirectory) = "" Then
        MsgBox "Root path not found. Please check Cell B2.", vbCritical
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    createdCount = 0

    On Error Resume Next
    For i = 5 To lastRow ' Data starts at Row 5
        folderName = ws.Cells(i, 1).Value
        If folderName <> "" Then
            If Dir(rootPath & folderName, vbDirectory) = "" Then
                MkDir rootPath & folderName
                createdCount = createdCount + 1
            End If
        End If
    Next i
    On Error GoTo 0

    MsgBox createdCount & " folders were successfully created.", vbInformation, "Process Complete"
End Sub
