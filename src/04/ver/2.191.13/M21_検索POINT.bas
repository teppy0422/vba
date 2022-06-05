Attribute VB_Name = "M21_����POINT"
Public sheetArray As Variant
Public �[�����c As Long, �[��c As Long, cavC As Long, LEDc As Long, �|�C���g1c As Long, �|�C���g2c As Long, fuseC As Long, ��d�W�~c As Long
Public LEDval As String, �|�C���g1val As String, �|�C���g2val As String, FUSEval As String, ��d�W�~val As String

Public Function POINTset(book, sheet)
    Dim maxRow As Long, maxCol As Long
    Dim pointTitle As Range
    With Workbooks(book).Sheets(sheet)
        maxCol = .Cells(2, 1).End(xlToRight).Column
        Set pointTitle = .Range(.Cells(2, 1), .Cells(2, maxCol))
        �[�����c = pointTitle.Find("�[�����i��", , , xlWhole).Column: If �[�����c = 0 Then Stop
        �[��c = pointTitle.Find("�[����", , , xlWhole).Column: If �[��c = 0 Then Stop
        cavC = pointTitle.Find("Cav", , , xlWhole).Column: If cavC = 0 Then Stop
        LEDc = pointTitle.Find("LED", , , xlWhole).Column: If LEDc = 0 Then Stop
        �|�C���g1c = pointTitle.Find("�|�C���g1", , , xlWhole).Column: If �|�C���g1c = 0 Then Stop
        �|�C���g2c = pointTitle.Find("�|�C���g2", , , xlWhole).Column: If �|�C���g2c = 0 Then Stop
        fuseC = pointTitle.Find("FUSE", , , xlWhole).Column: If fuseC = 0 Then Stop
        ��d�W�~c = pointTitle.Find("��d�W�~", , , xlWhole).Column: If ��d�W�~c = 0 Then Stop
        maxRow = .Cells(.Rows.count, �[��c).End(xlUp).Row
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
        If Replace(sheetArray(i, �[��c), " ", "") = Ter Then
            If Replace(sheetArray(i, �[�����c), " ", "") = partName Then
                If Replace(sheetArray(i, cavC), " ", "") = cav Then
                    LEDval = sheetArray(i, LEDc)
                    �|�C���g1val = sheetArray(i, �|�C���g1c)
                    �|�C���g2val = sheetArray(i, �|�C���g2c)
                    FUSEval = sheetArray(i, fuseC)
                    ��d�W�~val = sheetArray(i, ��d�W�~c)
                    found = 1
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

