Attribute VB_Name = "Module_createPDF"
Option Explicit

Sub createPDF()
    Dim lastRow As Long, lastCol As Long
    Dim dat As Variant
    Dim bRet As Boolean
    Dim sh As New clSheet
    bRet = sh.getDataAsArray(ThisWorkbook, Sheet_tool.Name, shTool.list_row, 0, shTool.no_col, shTool.outputPdf_col, dat, lastRow, lastCol)
    
    Dim fullPath As String, pdfFullPath As String
    Dim i As Long
    For i = 1 To UBound(dat) Step 1
        fullPath = dat(i, shTool.tgtPath_col) & "\" & dat(i, shTool.tgtExcel_col)
        
        Dim wb As Workbook, ws As Worksheet
        If Not IsEmpty(fullPath) And fullPath <> "" Then
            Set wb = Workbooks.Open(fullPath)
            For Each ws In wb.Worksheets
                With ws.PageSetup
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                    .CenterHorizontally = False
                    .PaperSize = xlPaperA4
                    .Orientation = xlPortrait
                End With
            Next ws
            Worksheets.Select
            pdfFullPath = dat(i, shTool.outputPath_col) & "\" & dat(i, shTool.outputPdf_col)
            wb.ExportAsFixedFormat Type:=xlTypePDF, _
                                                                fileName:=pdfFullPath, _
                                                                IgnorePrintAreas:=False, _
                                                                IncludeDocProperties:=True, _
                                                                Quality:=xlQualityStandard, _
                                                                OpenAfterPublish:=False
            
            Sheet_tool.Cells(shTool.list_row + i - 1, shTool.note_col).Value = "pdfèoóÕÇµÇ‹ÇµÇΩÅB"

            Application.DisplayAlerts = False
            wb.Close
            Application.DisplayAlerts = True
        End If
    Next i
    
    MsgBox "complete!"

End Sub


