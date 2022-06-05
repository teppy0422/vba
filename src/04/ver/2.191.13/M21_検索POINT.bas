Attribute VB_Name = "M21_検索POINT"
Public sheetArray As Variant
Public 端末矢崎c As Long, 端末c As Long, cavC As Long, LEDc As Long, ポイント1c As Long, ポイント2c As Long, fuseC As Long, 二重係止c As Long
Public LEDval As String, ポイント1val As String, ポイント2val As String, FUSEval As String, 二重係止val As String

Public Function POINTset(book, sheet)
    Dim maxRow As Long, maxCol As Long
    Dim pointTitle As Range
    With Workbooks(book).Sheets(sheet)
        maxCol = .Cells(2, 1).End(xlToRight).Column
        Set pointTitle = .Range(.Cells(2, 1), .Cells(2, maxCol))
        端末矢崎c = pointTitle.Find("端末矢崎品番", , , xlWhole).Column: If 端末矢崎c = 0 Then Stop
        端末c = pointTitle.Find("端末№", , , xlWhole).Column: If 端末c = 0 Then Stop
        cavC = pointTitle.Find("Cav", , , xlWhole).Column: If cavC = 0 Then Stop
        LEDc = pointTitle.Find("LED", , , xlWhole).Column: If LEDc = 0 Then Stop
        ポイント1c = pointTitle.Find("ポイント1", , , xlWhole).Column: If ポイント1c = 0 Then Stop
        ポイント2c = pointTitle.Find("ポイント2", , , xlWhole).Column: If ポイント2c = 0 Then Stop
        fuseC = pointTitle.Find("FUSE", , , xlWhole).Column: If fuseC = 0 Then Stop
        二重係止c = pointTitle.Find("二重係止", , , xlWhole).Column: If 二重係止c = 0 Then Stop
        maxRow = .Cells(.Rows.count, 端末c).End(xlUp).Row
        sheetArray = .Range(.Cells(1, 1), .Cells(maxRow, maxCol))
        'ResultC = pointTitle.Find("PVSWtoPOINT_", , , xlWhole).Column: If ResultC = 0 Then Stop
    End With
    Set pointTitle = Nothing
End Function

Public Function POINTrelease()
    sheetArray = ""
End Function

Public Function POINTseek(partName, Ter, cav, ByRef found)
    Dim i As Long
    found = 0
    For i = 2 To UBound(sheetArray)
        If Replace(sheetArray(i, 端末c), " ", "") = Ter Then
            If Replace(sheetArray(i, 端末矢崎c), " ", "") = partName Then
                If Replace(sheetArray(i, cavC), " ", "") = cav Then
                    LEDval = sheetArray(i, LEDc)
                    ポイント1val = sheetArray(i, ポイント1c)
                    ポイント2val = sheetArray(i, ポイント2c)
                    FUSEval = sheetArray(i, fuseC)
                    二重係止val = sheetArray(i, 二重係止c)
                    found = 1
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

