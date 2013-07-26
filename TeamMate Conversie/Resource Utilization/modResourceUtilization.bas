Attribute VB_Name = "modResourceUtilization"
Option Explicit

Public Sub Resource_Utilization()
    Dim oWkb As Workbook
    Dim oWksOut As Worksheet
    Dim oWksIn As Worksheet
    
    Dim iRowIn As Long
    Dim iRowOut As Long
    Dim aSplitStr As Variant
    
    Dim sNaam As String
    Dim sProject As String
    Dim bSkip As Boolean
    
    Application.ScreenUpdating = False
    
    Set oWkb = Application.ActiveWorkbook
    Set oWksIn = oWkb.Worksheets("Sheet1")
    
    Set oWksOut = Worksheets.Add()
    oWksOut.Name = "Uren"
   
    oWksOut.Cells(1, 1) = "Cd-Project"
    oWksOut.Cells(1, 2) = "Project"
    oWksOut.Cells(1, 3) = "Medewerker"
    oWksOut.Cells(1, 4) = "Uren"
    oWksOut.Cells(1, 5) = "Soort"
    
    iRowIn = 8
    iRowOut = 2
    bSkip = False
    
    oWksIn.Activate
    
    Do Until oWksIn.Cells(iRowIn, 1) = "Total"
        'Medewerkernaam
        If Trim(oWksIn.Cells(iRowIn, 1)) <> "" And InStr(Trim(oWksIn.Cells(iRowIn, 1)), "Resource") = 0 And InStr(Trim(oWksIn.Cells(iRowIn, 1)), "Generated") = 0 Then
            sNaam = Trim(oWksIn.Cells(iRowIn, 1))
        End If
        
        'Projectnaam
        If Trim(oWksIn.Cells(iRowIn, 2)) <> "" Then
            sProject = oWksIn.Cells(iRowIn, 2)
            
            'evt. laatste deel projectnaam
            If Trim(oWksIn.Cells(iRowIn + 1, 2)) <> "" And (Val(Trim(oWksIn.Cells(iRowIn + 1, 10))) = 0 And Val(Trim(oWksIn.Cells(iRowIn + 1, 14))) = 0 And Val(Trim(oWksIn.Cells(iRowIn + 1, 16))) = 0) Then
                sProject = sProject & oWksIn.Cells(iRowIn + 1, 2)
                bSkip = True
            End If

            If Val(Trim(oWksIn.Cells(iRowIn, 10))) <> 0 Then
                aSplitStr = Split(sProject, "|")
            Else
                aSplitStr = Split("x|y", "|")
                aSplitStr(0) = ""
                aSplitStr(1) = sProject
            End If
            
            oWksOut.Cells(iRowOut, 1) = aSplitStr(0)
            oWksOut.Cells(iRowOut, 2) = aSplitStr(1)
            oWksOut.Cells(iRowOut, 3) = sNaam
            
            If Val(Trim(oWksIn.Cells(iRowIn, 10))) <> 0 Then
                oWksOut.Cells(iRowOut, 4) = oWksIn.Cells(iRowIn, 10)
                oWksOut.Cells(iRowOut, 5) = "Project"
            ElseIf Val(Trim(oWksIn.Cells(iRowIn, 14))) <> 0 Then
                oWksOut.Cells(iRowOut, 4) = oWksIn.Cells(iRowIn, 14)
                oWksOut.Cells(iRowOut, 5) = "Nonworking"
            ElseIf Val(Trim(oWksIn.Cells(iRowIn, 16))) <> 0 Then
                oWksOut.Cells(iRowOut, 4) = oWksIn.Cells(iRowIn, 16)
                oWksOut.Cells(iRowOut, 5) = "Admin"
            End If

            iRowOut = iRowOut + 1
            
        End If

        iRowIn = iRowIn + 1
        If bSkip Then
            iRowIn = iRowIn + 1
            bSkip = False
        End If
    Loop
   
    Application.ScreenUpdating = True

    Set oWksIn = Nothing
    Set oWksOut = Nothing
    Set oWkb = Nothing
End Sub
