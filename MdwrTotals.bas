Attribute VB_Name = "MdwrTotals"
Option Explicit

Public Sub MdwrTotals()
    Dim oWkb As Workbook
    Dim oInpWks As Worksheet
    Dim oOutWeek As Worksheet
    Dim oOutMaand As Worksheet
    Dim oDump As Worksheet
    
    Dim iWeekStartCol As Integer
    Dim iWeekEndCol As Integer
    Dim iMaandStartCol As Integer
    Dim iMaandEndCol As Integer
    
    Dim iWeekRow As Long
    Dim iMaandRow As Long
    Dim iDumprow As Long
    
    Set oWkb = Application.ActiveWorkbook
    Set oOutWeek = tblWeek
    Set oOutMaand = tblMaand
    Set oDump = tblDump
                
    ClearSheet oOutWeek, 2, 1
    ClearSheet oOutMaand, 2, 1
    ClearSheet oDump, 2, 1
            
    Application.ScreenUpdating = False
            
    iWeekRow = 2
    iMaandRow = 2
    iDumprow = 2
            
    For Each oInpWks In oWkb.Worksheets
        If Application.WorksheetFunction.IsLogical(oInpWks.Cells(1, 2)) Then
            If oInpWks.CodeName <> "leegMedewerker" And oInpWks.CodeName <> "leegProject" Then
                If oInpWks.Cells(2, 2) <> "P" And oInpWks.Cells(2, 2) <> "M" Then
                    MsgBox "Geen geldig werkblad!", vbCritical, "Applicatiefout"
                    Exit Sub
                End If
                
                If Trim(oInpWks.Cells(2, 2)) = "P" Then
                    iWeekStartCol = 17
                    iWeekEndCol = 87
                    iMaandStartCol = 89
                    iMaandEndCol = 105
                    
                    iWeekRow = FillSheet(oInpWks, iWeekStartCol, iWeekEndCol, oOutWeek, iWeekRow)
                    iMaandRow = FillSheet(oInpWks, iMaandStartCol, iMaandEndCol, oOutMaand, iMaandRow)
                    iDumprow = DumpSheet(oInpWks, oDump, iDumprow)
                End If
            End If
        End If
    Next
    
    Application.ScreenUpdating = True
    
    Set oInpWks = Nothing
    Set oWkb = Nothing
End Sub

Private Sub ClearSheet(oWks As Worksheet, iStartRow As Long, iStartCol As Long)
    oWks.Activate
    oWks.Cells(iStartRow + 1, iStartCol + 1).Value = "XXX"
    oWks.Cells(iStartRow, iStartCol).Select
    oWks.Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    oWks.Cells(iStartRow, iStartCol).Select
End Sub

Private Function DumpSheet(oInpWks As Worksheet, oOutWks As Worksheet, iOutRow As Long)
    Dim iInpRow As Long
    
    oInpWks.Activate
    oInpWks.Cells(9, 3).Activate
    iInpRow = 9
    
    Do Until IsEmpty(oInpWks.Cells(iInpRow, 3))
        oOutWks.Cells(iOutRow, 1) = oInpWks.Cells(iInpRow, 10)
        oOutWks.Cells(iOutRow, 2) = oInpWks.Cells(iInpRow, 11)
        oOutWks.Cells(iOutRow, 3) = oInpWks.Cells(iInpRow, 12)
        oOutWks.Cells(iOutRow, 4) = oInpWks.Cells(iInpRow, 13)
        oOutWks.Cells(iOutRow, 5) = oInpWks.Cells(iInpRow, 14)
        oOutWks.Cells(iOutRow, 6) = oInpWks.Cells(iInpRow, 15)
        oOutWks.Cells(iOutRow, 7) = oInpWks.Cells(iInpRow, 6)
        
        iOutRow = iOutRow + 1
        iInpRow = iInpRow + 1
    Loop
    
    DumpSheet = iOutRow
End Function

Private Function FillSheet(oInpWks As Worksheet, iStartCol As Integer, iEndCol As Integer, oOutWks As Worksheet, iOutRow As Long) As Long
    Dim iInpRow As Long
    Dim iInpCol As Integer
    
    
    oInpWks.Activate
    oInpWks.Cells(9, 3).Activate
    iInpRow = 9
    
    Do Until IsEmpty(oInpWks.Cells(iInpRow, 3))
        For iInpCol = iStartCol To iEndCol
            If IsNumeric(oInpWks.Cells(iInpRow, iInpCol)) And oInpWks.Cells(iInpRow, iInpCol) <> 0 Then
                oOutWks.Cells(iOutRow, 1) = oInpWks.Cells(iInpRow, 10)
                oOutWks.Cells(iOutRow, 2) = oInpWks.Cells(iInpRow, 11)
                oOutWks.Cells(iOutRow, 3) = oInpWks.Cells(iInpRow, 12)
                oOutWks.Cells(iOutRow, 4) = oInpWks.Cells(iInpRow, 13)
                oOutWks.Cells(iOutRow, 5) = oInpWks.Cells(8, iInpCol)
                oOutWks.Cells(iOutRow, 6) = Year(oInpWks.Cells(6, iInpCol))
                oOutWks.Cells(iOutRow, 7) = oInpWks.Cells(iInpRow, iInpCol)
                iOutRow = iOutRow + 1
            End If
        Next
        
        iInpRow = iInpRow + 1
    Loop
    
    FillSheet = iOutRow
End Function


