Attribute VB_Name = "FileMover"
' ====================================================================
' Tool: Bulk File Mover
' Purpose: Transfers files from a Source Path to a Destination Path
' Developer: SG (Strategic Admin Automation)
' ====================================================================

Sub MoveFilesToNewFolder()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourcePath As String, destPath As String
    Dim fileName As String, successCount As Integer
    Dim fso As Object

    ' Use ActiveSheet to make it flexible for the user
    Set ws = ActiveSheet
    
    ' Late binding for FileSystemObject (no references needed)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Configuration: Adjust these cells based on your Excel layout
    sourcePath = Trim(ws.Range("B2").Value) ' Source Folder Path
    destPath = Trim(ws.Range("B3").Value)   ' Destination Folder Path
    
    ' Add trailing backslashes if missing
    If sourcePath <> "" And Right(sourcePath, 1) <> "\" Then sourcePath = sourcePath & "\"
    If destPath <> "" And Right(destPath, 1) <> "\" Then destPath = destPath & "\"

    ' 1. Validate Source and Destination Paths
    If Dir(sourcePath, vbDirectory) = "" Or Dir(destPath, vbDirectory) = "" Then
        MsgBox "Error: One or both folder paths are invalid. Please check B2 and B3.", vbCritical, "Path Error"
        Exit Sub
    End If

    ' Find the last row of filenames in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there is data to process
    If lastRow < 6 Then
        MsgBox "No filenames found. Please list files starting from Row 6.", vbExclamation
        Exit Sub
    End If

    successCount = 0

    ' 2. Loop through the filenames and move them
    For i = 6 To lastRow
        fileName = Trim(ws.Cells(i, 1).Value)
        
        If fileName <> "" Then
            ' Check if the file exists in the source folder
            If fso.FileExists(sourcePath & fileName) Then
                On Error Resume Next
                
                ' Perform the move
                fso.MoveFile Source:=sourcePath & fileName, Destination:=destPath & fileName
                
                ' Log status in Column B
                If Err.Number = 0 Then
                    ws.Cells(i, 2).Value = "Moved"
                    ws.Cells(i, 2).Font.Color = RGB(0, 128, 0) ' Green
                    successCount = successCount + 1
                Else
                    ws.Cells(i, 2).Value = "Error: " & Err.Description
                    ws.Cells(i, 2).Font.Color = vbRed
                End If
                On Error GoTo 0
            Else
                ws.Cells(i, 2).Value = "Not Found in Source"
                ws.Cells(i, 2).Font.Color = vbBlue
            End If
        End If
    Next i

    ' 3. Final Summary
    MsgBox successCount & " files were moved successfully!", vbInformation, "Task Complete"
End Sub