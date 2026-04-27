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

    Set ws = ThisWorkbook.Sheets("File_Renamer")
    fPath = ws.Range("B2").Value
    
    If Right(fPath, 1) <> "\" Then fPath = fPath & "\"

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    successCount = 0

    For i = 5 To lastRow
        oldName = ws.Cells(i, 1).Value
        newName = ws.Cells(i, 2).Value
        
        If oldName <> "" And newName <> "" Then
            If Dir(fPath & oldName) <> "" Then
                Name fPath & oldName As fPath & newName
                ws.Cells(i, 3).Value = "Renamed"
                successCount = successCount + 1
            Else
                ws.Cells(i, 3).Value = "Original File Not Found"
            End If
        End If
    Next i

    MsgBox successCount & " files renamed successfully.", vbInformation
End Sub
