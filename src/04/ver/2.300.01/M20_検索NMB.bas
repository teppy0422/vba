Attribute VB_Name = "M20_����NMB"
Public sheetArray As Variant
Public ���ic As Long, �\��c As Long, �i��c As Long, �T�C�Yc As Long, �Fc As Long, �i���c As Long, �T�C�Y��c As Long, �F��c As Long, ����c As Long, ����c As Long, JCDFc As Long, ���i�i��c As Long
Public ��1c As Long, ���i11c As Long, �[��1c As Long, getFelt1c As Long, ��2c As Long, ���i21c As Long, �[��2c As Long, getFelt2c As Long
Public ���ival As String, �\��val As String, �i��val As String, �T�C�Yval As String, �Fval As String, �i���val As String, �T�C�Y��val As String, �F��val As String, ����val As String, ����val As String, JCDFval As String, ���i�i��val As String
Public ��1val As String, ���i11val As String, �[��1val As String, getFelt1val As String, ��2val As String, ���i21val As String, �[��2val As String, getFelt2val As String

Public Function NMBset(book, sheet)
    Dim maxRow As Long, maxCol As Long
    Dim nmbTitle As Range
    With Workbooks(book).Sheets(sheet)
        maxRow = .Cells(.Rows.count, 1).End(xlUp).Row
        maxCol = .Cells(1, 1).End(xlToRight).Column
        sheetArray = .Range(.Cells(1, 1), .Cells(maxRow, maxCol))
        Set nmbTitle = .Range(.Cells(1, 1), .Cells(1, maxCol))
        '�d���p
        ���ic = nmbTitle.Find("���i", , , xlWhole).Column: If ���ic = 0 Then Stop
        �\��c = nmbTitle.Find("�\��", , , xlWhole).Column: If �\��c = 0 Then Stop
        �i��c = nmbTitle.Find("�i��", , , xlWhole).Column: If �i��c = 0 Then Stop
        �T�C�Yc = nmbTitle.Find("����", , , xlWhole).Column: If �T�C�Yc = 0 Then Stop
        �Fc = nmbTitle.Find("�F", , , xlWhole).Column: If �Fc = 0 Then Stop
        �i���c = nmbTitle.Find("�i��", , , xlWhole).Column: If �i���c = 0 Then Stop
        �T�C�Y��c = nmbTitle.Find("�T��", , , xlWhole).Column: If �T�C�Y��c = 0 Then Stop
        �F��c = nmbTitle.Find("�F��", , , xlWhole).Column: If �F��c = 0 Then Stop
        ����c = nmbTitle.Find("����", , , xlWhole).Column: If ����c = 0 Then Stop
        ����c = nmbTitle.Find("����", , , xlWhole).Column: If ����c = 0 Then Stop
        JCDFc = nmbTitle.Find("JCDF", , , xlWhole).Column: If JCDFc = 0 Then Stop
        '�d���[���p
        ��1c = nmbTitle.Find("��1", , , xlWhole).Column: If ��1c = 0 Then Stop
        ���i11c = nmbTitle.Find("���i11", , , xlWhole).Column: If ���i11c = 0 Then Stop
        �[��1c = nmbTitle.Find("�[��1", , , xlWhole).Column: If �[��1c = 0 Then Stop
        getFelt1c = nmbTitle.Find("ό�1", , , xlWhole).Column: If getFelt1c = 0 Then Stop
        ��2c = nmbTitle.Find("��2", , , xlWhole).Column: If ��2c = 0 Then Stop
        ���i21c = nmbTitle.Find("���i21", , , xlWhole).Column: If ���i21c = 0 Then Stop
        �[��2c = nmbTitle.Find("�[��2", , , xlWhole).Column: If �[��2c = 0 Then Stop
        getFelt2c = nmbTitle.Find("ό�2", , , xlWhole).Column: If getFelt2c = 0 Then Stop
    End With
    Set nmbTitle = Nothing
End Function

Public Function NMBrelease()
    sheetArray = ""
End Function

Public Function NMBseek_�d��(product, cons, ByRef found, z)
    Dim i As Long
    found = 0
    For i = 2 To UBound(sheetArray)
        If Replace(sheetArray(i, �\��c), " ", "") = cons Then
            If Replace(sheetArray(i, ���ic), " ", "") = product Then
                ���i�i��val = sheetArray(i, ���ic)
                �i��val = sheetArray(i, �i��c)
                �T�C�Yval = sheetArray(i, �T�C�Yc)
                �Fval = sheetArray(i, �Fc)
                �i���val = Replace(sheetArray(i, �i���c), " ", "")
                �T�C�Y��val = Replace(sheetArray(i, �T�C�Y��c), " ", "")
                �F��val = Replace(sheetArray(i, �F��c), " ", "")
                ����val = sheetArray(i, ����c)
                ����val = Replace(sheetArray(i, ����c), " ", "")
                JCDFval = sheetArray(i, JCDFc)
                found = 1
                Exit Function
            End If
        End If
    Next i
End Function

Public Function NMBseek_�d���[��(product, cons, ByRef found)
    Dim i As Long
    found = 0
    For i = 2 To UBound(sheetArray)
        If Replace(sheetArray(i, �\��c), " ", "") = cons Then
            If Replace(sheetArray(i, ���ic), " ", "") = product Then
                �i��val = sheetArray(i, �i��c)
                �T�C�Yval = sheetArray(i, �T�C�Yc)
                �Fval = sheetArray(i, �Fc)
                �i���val = Replace(sheetArray(i, �i���c), " ", "")
                �T�C�Y��val = Replace(sheetArray(i, �T�C�Y��c), " ", "")
                �F��val = Replace(sheetArray(i, �F��c), " ", "")
                ����val = sheetArray(i, ����c)
                '�d���[���p
                ��1val = sheetArray(i, ��1c)
                ���i11val = sheetArray(i, ���i11c)
                �[��1val = sheetArray(i, �[��1c)
                getFelt1val = sheetArray(i, getFelt1c)
                ��2val = sheetArray(i, ��2c)
                ���i21val = sheetArray(i, ���i21c)
                �[��2val = sheetArray(i, �[��2c)
                getFelt2val = sheetArray(i, getFelt2c)
                found = "1"
                Exit Function
            End If
        End If
    Next i
End Function
