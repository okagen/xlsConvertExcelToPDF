Attribute VB_Name = "Module_createList"
Option Explicit
'------------------------------------------------------------
' Create file list.
'------------------------------------------------------------
Public Function createList(ByVal flPath As String, ByRef r As Long) As Long
    Dim buf As String, f As Object
    Dim flPath_show As String
    
    ' Set current path to flPath if the parent directory is empty.
    If flPath = "" Then
        flPath = ThisWorkbook.path
    End If
    
    ' Get all excel file.
    buf = Dir(flPath & "\*.xlsx")

    With Sheet_tool
        .Activate
        Do While buf <> ""
            r = r + 1
            .Cells(r, shTool.no_col).Value = r - shTool.list_row + 1
            .Cells(r, shTool.tgtPath_col).Value = flPath
            .Cells(r, shTool.tgtExcel_col).Value = buf
            .Cells(r, shTool.outputPath_col).Value = ThisWorkbook.path & "\" & DIR_PDF
            .Cells(r, shTool.outputPdf_col).Value = Replace(buf, "xlsx", "pdf")
            buf = Dir()
        Loop
    End With
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(flPath).SubFolders
            Call createList(f.path, r)
        Next f
    End With
    createList = r
End Function

'------------------------------------------------------------
' Clear data in the tool sheet.
'------------------------------------------------------------
Public Function clearSheet() As Boolean
    With Sheet_tool
        .Activate
        .Range(.Cells(shTool.list_row, shTool.no_col), _
                    .Cells(Range(.Cells(shTool.list_row, shTool.no_col), .Cells(shTool.list_row, shTool.no_col)).End(xlDown).row, shTool.no_col)).Select
        Selection.EntireRow.Delete
    End With
    clearSheet = True
End Function
        
'------------------------------------------------------------
' Set format.
'------------------------------------------------------------
Public Function setFormat(ByVal r As Long) As Boolean
    With Sheet_tool
        .Activate
        With .Range(.Cells(shTool.list_row, shTool.no_col), .Cells(r, shTool.outputPdf_col))
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            
        End With
    End With
    setFormat = True
End Function

