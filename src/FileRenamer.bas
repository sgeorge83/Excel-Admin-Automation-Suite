Attribute VB_Name = "FileRenamer"
' ====================================================================
' Tool: File Renamer & Archive
' Purpose: Batch renames files based on Excel mapping
' ====================================================================
Sub RenameFiles()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim fPath As String, oldName As String, newName As String
    Dim successCount As Integer
    Dim fullOldPath As String, fullNewPath As String

    Set ws = ActiveSheet
    fPath = Trim(ws.Range("B2").Value)
    
    ' 1. Fix the Backslash issue
    If fPath <> "" Then
        If Right(fPath, 1) <> "\" Then fPath = fPath & "\"
    End If

    ' 2. Check if Path is actually valid
    On Error Resume Next
    If Dir(fPath, vbDirectory) = "" Then
        MsgBox "Cannot find folder: " & fPath, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    successCount = 0

    For i = 5 To lastRow
        oldName = Trim(ws.Cells(i, 1).Value)
        newName = Trim(ws.Cells(i, 2).Value)
        
        If oldName <> "" And newName <> "" Then
            fullOldPath = fPath & oldName
            fullNewPath = fPath & newName
            
            ' 3. Use a more stable check for file existence
            If Len(Dir(fullOldPath)) > 0 Then
                On Error Resume Next
                Name fullOldPath As fullNewPath
                
                If Err.Number = 0 Then
                    ws.Cells(i, 3).Value = "Done"
                    ws.Cells(i, 3).Interior.Color = vbGreen
                    successCount = successCount + 1
                Else
                    ws.Cells(i, 3).Value = "Error: " & Err.Description
                    ws.Cells(i, 3).Interior.Color = vbRed
                End If
                On Error GoTo 0
            Else
                ws.Cells(i, 3).Value = "File Not Found"
                ws.Cells(i, 3).Interior.Color = vbYellow
            End If
        End If
    Next i

    MsgBox successCount & " files processed.", vbInformation
End Sub
