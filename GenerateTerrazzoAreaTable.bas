Attribute VB_Name = "GenerateTerrazzoAreaTable"
Sub GenerateTerrazzoAreaTable()
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    Dim ms As AcadModelSpace
    Set ms = doc.ModelSpace
    

    Dim ss As AcadSelectionSet
    On Error Resume Next
    Set ss = doc.SelectionSets("TempSelSet")
    If Not ss Is Nothing Then ss.Delete
    Set ss = doc.SelectionSets.Add("TempSelSet")
    On Error GoTo 0

    ss.SelectOnScreen

    If ss.count = 0 Then
        MsgBox "No text or polylines selected.", vbExclamation
        Exit Sub
    End If

    ' Separate collections
    Dim txts As Collection, pls As Collection
    Set txts = New Collection
    Set pls = New Collection

    Dim ent As AcadEntity
    For Each ent In ss
        If TypeOf ent Is AcadText Then
            txts.Add ent
        ElseIf TypeOf ent Is AcadLWPolyline Then
            Dim pl As AcadLWPolyline
            Set pl = ent
            If pl.Closed Then pls.Add pl
        End If
    Next

    If txts.count = 0 Or pls.count = 0 Then
        MsgBox "Selection must include at least one text and one closed polyline.", vbExclamation
        Exit Sub
    End If

    ' Match texts to containing polylines
    Dim result As Collection
    Set result = New Collection

    Dim txt As AcadText
    For Each txt In txts
        For Each pl In pls
            If PointInPolyline(pl, txt.InsertionPoint) Then
                Dim entry(1) As Variant
                entry(0) = txt.TextString
                entry(1) = pl.area
                result.Add entry
                
                ' Color both matched text and polyline green
                txt.color = acGreen
                pl.color = acGreen
                
                Exit For
            End If
        Next
    Next

    If result.count = 0 Then
        MsgBox "No matching text inside polylines found.", vbExclamation
        Exit Sub
    End If

    ' Insert table near first text insertion point
    Dim basePt(0 To 2) As Double
    Dim firstTxt As AcadText
    Set firstTxt = txts(1)
    
    basePt(0) = firstTxt.InsertionPoint(0) + 100
    basePt(1) = firstTxt.InsertionPoint(1)
    basePt(2) = 0

    Dim rowCount As Integer: rowCount = result.count + 1
    Dim colCount As Integer: colCount = 2

    Dim tableObj As AcadTable
    Set tableObj = ms.AddTable(basePt, rowCount, colCount, 10#, 50#)

    tableObj.SetText 0, 0, "Plot No."
    tableObj.SetText 0, 1, "Area (sq.units)"

    Dim i As Integer
    For i = 1 To result.count
        r = result(i)
        tableObj.SetText i, 0, r(0)
        tableObj.SetText i, 1, Format(r(1), "0.00")
    Next

    tableObj.Update
    MsgBox "Table inserted with " & result.count & " entries.", vbInformation
    
    '=== Export to Excel ===
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object

    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")  ' Try to hook to running Excel
    If Err.Number <> 0 Then
        Err.Clear
        Set xlApp = CreateObject("Excel.Application")  ' Start new Excel
    End If
    On Error GoTo 0

    If xlApp Is Nothing Then
        MsgBox "Excel could not be started.", vbExclamation
        Exit Sub
    End If

    xlApp.Visible = True
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Sheets(1)

    ' Write headers
    xlSheet.Cells(1, 1).Value = "Plot No."
    xlSheet.Cells(1, 2).Value = "Area (sq.units)"

    ' Write data
    Dim iRow As Integer
    For iRow = 1 To result.count
        r = result(iRow)
        xlSheet.Cells(iRow + 1, 1).Value = r(0)
        xlSheet.Cells(iRow + 1, 2).Value = r(1)
    Next

    MsgBox "Exported to Excel successfully.", vbInformation

End Sub


Function PointInPolyline(pl As AcadLWPolyline, pt As Variant) As Boolean
    Dim res As Boolean: res = False
    Dim coords As Variant: coords = pl.Coordinates
    Dim testX As Double: testX = pt(0)
    Dim testY As Double: testY = pt(1)
    Dim n As Integer: n = UBound(coords) \ 2

    Dim i As Integer, j As Integer
    For i = 0 To n
        j = (i + 1) Mod (n + 1)
        Dim xi As Double: xi = coords(2 * i)
        Dim yi As Double: yi = coords(2 * i + 1)
        Dim xj As Double: xj = coords(2 * j)
        Dim yj As Double: yj = coords(2 * j + 1)

        If ((yi > testY) <> (yj > testY)) Then
            Dim xInt As Double
            xInt = (xj - xi) * (testY - yi) / (yj - yi) + xi
            If testX < xInt Then res = Not res
        End If
    Next

    PointInPolyline = res
End Function

