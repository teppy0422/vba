Attribute VB_Name = "Module_Array"
Option Explicit

'�Q�����z�񂩂�l���擾�A�t�B�[���h�ԍ����w��A��؂蕶���̓J���}�B
'���ʂ���������ꍇ�͑Ή����ĂȂ�
Public Function searchRan(ByVal myRan As Variant, ByVal searchWords As String, ByVal searchFields As String, ByVal getFields As String) As String

    Dim searchWord As Variant, searchField  As Variant, getFieldNum As Variant, getField As Variant
    searchWord = Split(searchWords, ",")
    searchField = Split(searchFields, ",")
    getField = Split(getFields, ",")
    Dim i As Long, flag As Boolean, i_Anser As Long, resultTemp As String, x As Long
    
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        i_Anser = -1
        resultTemp = ""
        For x = LBound(searchField) To UBound(searchField)
            If searchWord(x) = myRan(searchField(x), i) Then
                If x = UBound(searchField) Then
                    i_Anser = i
                End If
            End If
        Next x
        If i_Anser <> -1 Then
            For x = LBound(getField) To UBound(getField)
                resultTemp = resultTemp & "," & �[�����i�ԕϊ�(myRan(getField(x), i_Anser))
            Next x
            searchRan = Mid(resultTemp, 2)
            Exit Function
        End If
    Next i
    searchRan = False
End Function
'�Q�����z�񂩂�l���擾�A�t�B�[���h�ԍ����w��A��؂蕶���̓J���}�B
'���ʂ���������ꍇ�͑Ή����ĂȂ�
Public Function searchRan_ver2(ByVal myRan As Variant, ByVal searchWords As String, ByVal searchFields As String, ByVal getFields As String) As String

    Dim searchWord As Variant, searchField  As Variant, getFieldNum As Variant, getField As Variant, x As Long, xx As Long
    searchWord = Split(searchWords, ",")
    searchField = Split(searchFields, ",")
    getField = Split(getFields, ",")
    Dim f As Variant
    ReDim f(UBound(searchField))
    '�����t�B�[���h�ԍ��̎擾
    For x = LBound(searchField) To UBound(searchField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, 1) = searchField(x) Then
                f(x) = xx
                GoTo line10
            End If
        Next xx
line10:
    Next x
    Dim g As Variant
    ReDim g(UBound(searchField))
    '�擾�t�B�[���h�ԍ��̎擾
    For x = LBound(getField) To UBound(getField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, 1) = getField(x) Then
                g(x) = xx
                GoTo line20
            End If
        Next xx
line20:
    Next x
    Dim i As Long, flag As Boolean, i_Anser As Long, resultTemp As String
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        i_Anser = -1
        resultTemp = ""
        For x = LBound(f) To UBound(f)
            If searchWord(x) = myRan(f(x), i) Then
                If x = UBound(f) Then
                    i_Anser = i
                    Exit For
                End If
            Else
                Exit For
            End If
        Next x
        If i_Anser <> -1 Then
            For x = LBound(getField) To UBound(getField)
                resultTemp = resultTemp & "," & myRan(g(x), i_Anser)
            Next x
            searchRan_ver2 = Mid(resultTemp, 2)
            Exit Function
        End If
    Next i
    searchRan_ver2 = False
End Function

'�Q�����z�񂩂�l���擾�A�t�B�[���h�ԍ����w��A��؂蕶���̓J���}�B
'���ʂ���������ꍇ�͑Ή����ĂȂ�
Public Function searchRan_ver3(ByVal myRan As Variant, ByVal searchWords As String, ByVal searchFields As String, ByVal getFields As String) As String

    Dim searchWord As Variant, searchField  As Variant, getFieldNum As Variant, getField As Variant, x As Long, xx As Long
    searchWord = Split(searchWords, ",")
    searchField = Split(searchFields, ",")
    getField = Split(getFields, ",")
    Dim f As Variant
    ReDim f(UBound(searchField))
    '�����t�B�[���h�ԍ��̎擾
    For x = LBound(searchField) To UBound(searchField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, 0) = searchField(x) Then
                f(x) = xx
                GoTo line10
            End If
        Next xx
line10:
    Next x
    Dim g As Variant
    ReDim g(UBound(searchField))
    '�擾�t�B�[���h�ԍ��̎擾
    For x = LBound(getField) To UBound(getField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, 0) = getField(x) Then
                g(x) = xx
                GoTo line20
            End If
        Next xx
line20:
    Next x
    Dim i As Long, flag As Boolean, i_Anser As Long, resultTemp As String
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        i_Anser = -1
        resultTemp = ""
        For x = LBound(f) To UBound(f)
            If searchWord(x) = myRan(f(x), i) Then
                If x = UBound(f) Then
                    i_Anser = i
                    Exit For
                End If
            Else
                Exit For
            End If
        Next x
        If i_Anser <> -1 Then
            For x = LBound(getField) To UBound(getField)
                resultTemp = resultTemp & "," & myRan(g(x), i_Anser)
            Next x
            searchRan_ver3 = Mid(resultTemp, 2)
            Exit Function
        End If
    Next i
    searchRan_ver3 = False
End Function

Public Function searchRan_ver4(ByVal myRan As Variant, ByVal searchWords As String, ByVal searchFields As String, ByVal getFields As String, ByVal searchRow As Integer) As String
    If searchWords = "" Then searchRan_ver4 = False: Exit Function
    
    Dim searchWord As Variant, searchField  As Variant, getFieldNum As Variant, getField As Variant, x As Long, xx As Long
    searchWord = Split(searchWords, ",")
    searchField = Split(searchFields, ",")
    getField = Split(getFields, ",")
    Dim f As Variant
    ReDim f(UBound(searchField))
    '�����t�B�[���h�ԍ��̎擾
    For x = LBound(searchField) To UBound(searchField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, searchRow) = searchField(x) Then
                f(x) = xx
                GoTo line10
            End If
        Next xx
line10:
    Next x
    Dim g As Variant
    ReDim g(UBound(getField))
    '�擾�t�B�[���h�ԍ��̎擾
    For x = LBound(getField) To UBound(getField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, searchRow) = getField(x) Then
                g(x) = xx
                GoTo line20
            End If
        Next xx
line20:
    Next x
    
    Dim i As Long, flag As Boolean, i_Anser As Long, resultTemp As String
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        i_Anser = -1
        resultTemp = ""
        For x = LBound(f) To UBound(f)
            If searchWord(x) = myRan(f(x), i) Then
                If x = UBound(f) Then
                    i_Anser = i
                    Exit For
                End If
            Else
                Exit For
            End If
        Next x
        If i_Anser <> -1 Then
            For x = LBound(getField) To UBound(getField)
                If IsEmpty(g(x)) Then
                    resultTemp = resultTemp & "," & "False"
                Else
                    resultTemp = resultTemp & "," & myRan(g(x), i_Anser)
                End If
            Next x
            searchRan_ver4 = Mid(resultTemp, 2)
            Exit Function
        End If
    Next i
    searchRan_ver4 = False
End Function

Public Function searchRan_xy(ByVal myRan As Variant, ByVal searchWords As String, ByVal searchFields As String, ByVal getFields As String, ByVal searchRow As Integer) As String
    If searchWords = "" Then searchRan_xy = False: Exit Function
    
    Dim searchWord As Variant, searchField  As Variant, getFieldNum As Variant, getField As Variant, x As Long, xx As Long
    searchWord = Split(searchWords, ",")
    searchField = Split(searchFields, ",")
    getField = Split(getFields, ",")
    Dim f As Variant
    ReDim f(UBound(searchField))
    '�����t�B�[���h�ԍ��̎擾
    For x = LBound(searchField) To UBound(searchField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, searchRow) = searchField(x) Then
                f(x) = xx
                GoTo line10
            End If
        Next xx
line10:
    Next x
    Dim g As Variant
    ReDim g(UBound(getField))
    '�擾�t�B�[���h�ԍ��̎擾
    For x = LBound(getField) To UBound(getField)
        For xx = LBound(myRan) To UBound(myRan)
            If myRan(xx, searchRow) = getField(x) Then
                g(x) = xx
                GoTo line20
            End If
        Next xx
line20:
    Next x
    
    Dim i As Long, flag As Boolean, i_Anser As Long, resultTemp As String
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        i_Anser = -1
        resultTemp = ""
        For x = LBound(f) To UBound(f)
            If searchWord(x) = myRan(f(x), i) Then
                If x = UBound(f) Then
                    i_Anser = i
                    Exit For
                End If
            Else
                Exit For
            End If
        Next x
        If i_Anser <> -1 Then
            For x = LBound(getField) To UBound(getField)
                If IsEmpty(g(x)) Then
                    resultTemp = resultTemp & "," & "False"
                Else
                    resultTemp = resultTemp & "," & g(x) & "," & i_Anser
                End If
            Next x
            
            searchRan_xy = Mid(resultTemp, 2)
            Exit Function
        End If
    Next i
    searchRan_xy = False
End Function

'key���u�����N�Ȃ�Z�b�g���Ȃ�,
Function readSheetToRan(ByVal mySheet As Worksheet, ByVal keyWord As Variant, ByVal FieldNames As String, ByVal blankKey As String) As Variant
    Dim key As Variant, sp As Variant, fieldname As String, i As Long, f As Variant, lastRow As Long, C As Long, x As Long, blankKeyPos As Long
    sp = Split(FieldNames, ",")
    ReDim f(UBound(sp))
    Dim myRan() As Variant
    ReDim myRan(UBound(sp), 0)
    With mySheet
        Set key = .Cells.Find(���i�i��str, , , 1)
        For i = LBound(sp) To UBound(sp)
            f(i) = .Rows(key.Row).Find(sp(i), , , 1).Column
            myRan(i, 0) = sp(i)
            If sp(i) = blankKey Then blankKeyPos = i
        Next i
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        For i = key.Row + 1 To lastRow
            If .Cells(i, key.Column) <> "" Then
                ReDim Preserve myRan(UBound(sp), UBound(myRan, 2) + 1)
                For x = LBound(sp) To UBound(sp)
                    myRan(x, UBound(myRan, 2)) = .Cells(i, f(x))
                Next x
            End If
        Next i
    End With
    
    '�u�����N�̏ꍇ�̓Z�b�g���Ȃ�
    Dim myRan2() As Variant
    ReDim myRan2(UBound(sp), 0)
    '�s�v�ȍs����菜��
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        If myRan(blankKeyPos, i) <> "" Then
            For x = LBound(sp) To UBound(sp)
                myRan2(x, UBound(myRan2, 2)) = myRan(x, i)
            Next x
            ReDim Preserve myRan2(UBound(sp), UBound(myRan2, 2) + 1)
        End If
    Next i
    If UBound(myRan2, 2) = 0 Then Stop '���i���X�g�̍H��a��40����͂��ꂽ�`���[�u������
    ReDim Preserve myRan2(UBound(sp), UBound(myRan2, 2) - 1)
    
    readSheetToRan = myRan2
End Function

'key���u�����N�Ȃ�Z�b�g���Ȃ�
'dataType=1:value,2:interior.color,3:font.color
Function readSheetToRan3(ByVal mySheet As Worksheet, ByVal keyWord As Variant, ByVal FieldNames As String, _
    ByVal blankKey As String, Optional ByVal dataType As Integer) As Variant

    Dim key As Variant, sp As Variant, fieldname As String, i As Long, f As Variant, _
        lastRow As Long, C As Long, x As Long, blankKeyPos As Long
        
    sp = Split(FieldNames, ",")
    ReDim f(UBound(sp))
    Dim myRan() As Variant
    ReDim myRan(UBound(sp), 0)
    With mySheet
        Set key = .Cells.Find(keyWord, , , 1)
        For i = LBound(sp) To UBound(sp)
            f(i) = .Rows(key.Row).Find(sp(i), , , 1, , , 1).Column
            myRan(i, 0) = sp(i)
            If sp(i) = blankKey Then blankKeyPos = i
        Next i
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        For i = key.Row + 1 To lastRow
            If .Cells(i, key.Column) <> "" Then
                ReDim Preserve myRan(UBound(sp), UBound(myRan, 2) + 1)
                For x = LBound(sp) To UBound(sp)
                    Select Case dataType
                        Case 1
                        myRan(x, UBound(myRan, 2)) = .Cells(i, f(x)).Value
                        Case 2
                        myRan(x, UBound(myRan, 2)) = .Cells(i, f(x)).Interior.color
                        Case 3
                        myRan(x, UBound(myRan, 2)) = .Cells(i, f(x)).Font.color
                        Case Else
                        Stop '�p�ӂ��Ă��Ȃ�
                    End Select
                Next x
            End If
        Next i
    End With
    
    '�u�����N�̏ꍇ�̓Z�b�g���Ȃ�
    Dim myRan2() As Variant
    ReDim myRan2(UBound(sp), 0)
    '�s�v�ȍs����菜��
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        If myRan(blankKeyPos, i) <> "" Then
            For x = LBound(sp) To UBound(sp)
                myRan2(x, UBound(myRan2, 2)) = myRan(x, i)
            Next x
            ReDim Preserve myRan2(UBound(sp), UBound(myRan2, 2) + 1)
        End If
    Next i
    If UBound(myRan2, 2) = 0 Then Stop '���i���X�g�̍H��a��40����͂��ꂽ�`���[�u������
    ReDim Preserve myRan2(UBound(sp), UBound(myRan2, 2) - 1)
    
    readSheetToRan3 = myRan2
End Function
'key���u�����N�Ȃ�Z�b�g���Ȃ�
Function readSheetToRan2(ByVal mySheet As Worksheet, ByVal keyWord As Variant, ByVal FieldNames As String, ByVal blankKey As String) As Variant

    Dim key As Variant, sp As Variant, fieldname As String, i As Long, f As Variant, _
        lastRow As Long, C As Long, x As Long, blankKeyPos As Long
        
    sp = Split(FieldNames, ",")
    ReDim f(UBound(sp))
    Dim myRan() As Variant
    ReDim myRan(UBound(sp), 0)
    With mySheet
        Set key = .Cells.Find(keyWord, , , 1)
        For i = LBound(sp) To UBound(sp)
            f(i) = .Rows(key.Row).Find(sp(i), , , 1, , , 1).Column
            myRan(i, 0) = sp(i)
            If sp(i) = blankKey Then blankKeyPos = i
        Next i
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        For i = key.Row + 1 To lastRow
            If .Cells(i, key.Column) <> "" Then
                ReDim Preserve myRan(UBound(sp), UBound(myRan, 2) + 1)
                For x = LBound(sp) To UBound(sp)
                    myRan(x, UBound(myRan, 2)) = .Cells(i, f(x))
                Next x
            End If
        Next i
    End With
    
    '�u�����N�̏ꍇ�̓Z�b�g���Ȃ�
    Dim myRan2() As Variant
    ReDim myRan2(UBound(sp), 0)
    '�s�v�ȍs����菜��
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        If myRan(blankKeyPos, i) <> "" Then
            For x = LBound(sp) To UBound(sp)
                myRan2(x, UBound(myRan2, 2)) = myRan(x, i)
            Next x
            ReDim Preserve myRan2(UBound(sp), UBound(myRan2, 2) + 1)
        End If
    Next i
    If UBound(myRan2, 2) = 0 Then Stop '���i���X�g�̍H��a��40����͂��ꂽ�`���[�u������
    ReDim Preserve myRan2(UBound(sp), UBound(myRan2, 2) - 1)
    
    readSheetToRan2 = myRan2
End Function

Function delete_Ran(ByVal myRan As Variant, ByVal deleteWord As String, ByVal deleteColumn As Long) As Variant
    Dim i As Long, myRan2() As Variant, x As Long
    ReDim myRan2(UBound(myRan), 0)
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        If myRan(deleteColumn, i) <> deleteWord And myRan(deleteColumn, i) <> "" Then
            For x = LBound(myRan) To UBound(myRan)
                myRan2(x, UBound(myRan2, 2)) = myRan(x, i)
            Next x
            ReDim Preserve myRan2(UBound(myRan), UBound(myRan2, 2) + 1)
        End If
    Next i
    ReDim Preserve myRan2(UBound(myRan), UBound(myRan2, 2) - 1)
    delete_Ran = myRan2
End Function

Function delete_RanVer2(ByVal myRan As Variant, ByVal deleteI As Long) As Variant
    Dim i As Long, x As Long, myRan2() As Variant
    ReDim myRan2(UBound(myRan), 0)
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        If i <> deleteI Then
            For x = LBound(myRan) To UBound(myRan)
                myRan2(x, UBound(myRan2, 2)) = myRan(x, i)
            Next x
            ReDim Preserve myRan2(UBound(myRan), UBound(myRan2, 2) + 1)
        End If
    Next i
    ReDim Preserve myRan2(UBound(myRan), UBound(myRan2, 2) - 1)
    delete_RanVer2 = myRan2
End Function

'��Ԗڂ̗v�f�̐��͕ύX�ł��Ȃ�����K�v
Function removeArrayIndex(ByVal targetArray As Variant, ByVal deleteIndex As Long) As Variant
    Dim i As Long, addRow As Long, x As Long
    Dim myRan() As Variant
    ReDim myRan(UBound(targetArray) - 1, UBound(targetArray, 2))
    '�폜�������v�f�ȊO��V�����z��ɓ����
    addRow = LBound(targetArray)
    For i = LBound(targetArray) To UBound(targetArray)
        If i <> deleteIndex Then
            For x = LBound(targetArray, 2) To UBound(targetArray, 2)
                myRan(addRow, x) = targetArray(i, x)
            Next x
            addRow = addRow + 1
        End If
    Next i
    removeArrayIndex = myRan
End Function

Sub export_ArrayToSheet(ByVal myArray As Variant, ByVal sheetName As String, ByVal transpose As Boolean)
    
    If transpose Then myArray = WorksheetFunction.transpose(myArray)
    
    If check_Sheet_existence(sheetName, wb(0)) = False Then
        Dim newWs As Worksheet
        Set newWs = Worksheets.add(after:=Worksheets("PVSW_RLTF"))
        newWs.Name = sheetName
        newWs.Cells.NumberFormat = "@"
    End If
    With wb(0).Sheets(sheetName)
        .Cells.ClearContents
        .Range(.Cells(1, 1), .Cells(UBound(myArray) + 1, UBound(myArray, 2) + 1)) = myArray
    End With
    
End Sub

Sub export_Array_ShiftJis(ByVal ran As Variant, ByVal myPath As String, ByVal Delimiter As String)
    '�e�L�X�g�t�@�C���ɂ��ďo��
    Dim lntFlNo As Integer: lntFlNo = FreeFile
    Dim outPutAddress As String: outPutAddress = myPath
    Open outPutAddress For Output As #lntFlNo
    Dim myLine As Variant, subSubNumber As String, myNow As Date, w As Variant, i As Long, ii As Long
    w = Chr(34)
    myNow = now
    For i = LBound(ran) To UBound(ran)
        myLine = ""
        For ii = LBound(ran, 2) To UBound(ran, 2)
                myLine = myLine & ran(i, ii) & Chr(44)
        Next ii
        myLine = myLine & myNow
        Print #lntFlNo, myLine
    Next i
    
    Close lntFlNo
End Sub

Sub export_Array_ShiftJis_ver2(ByVal ran As Variant, ByVal myPath As String, ByVal Delimiter As String)
    '�e�L�X�g�t�@�C���ɂ��ďo��
    Dim lntFlNo As Integer: lntFlNo = FreeFile
    Dim outPutAddress As String: outPutAddress = myPath
    Open outPutAddress For Output As #lntFlNo
    Dim myLine As Variant, subSubNumber As String, myNow As Date, w As Variant, i As Long, ii As Long
    For i = LBound(ran, 2) To UBound(ran, 2)
        myLine = ""
        For ii = LBound(ran, 1) To UBound(ran, 1)
                myLine = myLine & Delimiter & ran(ii, i)
        Next ii
        myLine = Mid(myLine, 2)
        Print #lntFlNo, myLine
    Next i
    Close lntFlNo
End Sub

Function merge_Array(ByVal Array1 As Variant, ByVal Array2 As Variant) As Variant
    
    If UBound(Array1, 2) <> UBound(Array2, 2) Then Stop '�z��̒������قȂ�ꍇ�}�[�W���Ȃ�
    
    Dim mergeArray() As Variant, x As Long, i As Long
    ReDim mergeArray(UBound(Array1) + UBound(Array2) + 1, UBound(Array1, 2))
    
    Dim add As Long
    add = UBound(Array1)
    For i = LBound(mergeArray, 2) To UBound(mergeArray, 2)
        For x = LBound(Array1) To UBound(Array1)
            mergeArray(x, i) = Array1(x, i)
        Next x
        For x = LBound(Array2) To UBound(Array2)
            mergeArray(x + add + 1, i) = Array2(x, i)
        Next x
    Next i
    
    merge_Array = mergeArray
End Function

Public Function Sort_Array_select(ByRef ary As Variant, ByVal FieldNames As String, ByVal is_Ascendings As String)

    Dim i As Integer, j As Integer
    Dim x As Integer, xx As Long
    Dim Swap
    
    Dim k As Integer
    
    Dim fieldNameSp As Variant, is_AscendingSP As Variant, Xs As Variant
    fieldNameSp = Split(FieldNames, ",")
    is_AscendingSP = Split(is_Ascendings, ",")
    If UBound(fieldNameSp) <> UBound(is_AscendingSP) Then Stop '�v�f��������Ȃ�
    ReDim Xs(UBound(fieldNameSp))
    
    '�t�B�[���h���̈ʒu���m�F
    For i = LBound(fieldNameSp) To UBound(fieldNameSp)
        For x = LBound(ary) To UBound(ary)
            If fieldNameSp(i) = ary(x, 0) Then
                Xs(i) = x
                Exit For
            End If
        Next x
    Next i
    
    For i = LBound(ary, 2) + 1 To UBound(ary, 2)
        For j = LBound(ary, 2) + 1 To UBound(ary, 2)
            If i <> j Then
                For x = 0 To UBound(Xs)
                    If is_AscendingSP(x) Then
                        If compare_Text(ary(Xs(x), j), ary(Xs(x), i)) Then
                            For xx = LBound(ary) To UBound(ary)
                                Swap = ary(xx, i)
                                ary(xx, i) = ary(xx, j)
                                ary(xx, j) = Swap
                            Next xx
                        ElseIf ary(Xs(x), i) = ary(Xs(x), j) Then
                            GoTo line20 '�l���������������̃t�B�[���h�̊m�F�ɐi��
                        End If
                    Else
                        If compare_Text(ary(Xs(x), i), ary(Xs(x), j)) Then
                            For xx = LBound(ary) To UBound(ary)
                                Swap = ary(xx, i)
                                ary(xx, i) = ary(xx, j)
                                ary(xx, j) = Swap
                            Next xx
                        ElseIf ary(Xs(x), i) = ary(Xs(x), j) Then
                            GoTo line20 '�l���������������̃t�B�[���h�̊m�F�ɐi��
                        End If
                    End If
                    Exit For
line20:
                Next x
            End If
        Next j
    Next i
    
 End Function
Function mix_Array(ByVal baseRan As Variant, ByVal addRan As Variant) As Variant
    Dim i As Long, x As Long
    For i = LBound(addRan, 2) To UBound(addRan, 2)
        ReDim Preserve baseRan(UBound(baseRan), UBound(baseRan, 2) + 1)
        For x = LBound(baseRan) To UBound(baseRan)
            baseRan(x, UBound(baseRan, 2)) = addRan(x, i)
        Next x
    Next i
    mix_Array = baseRan
End Function

Function readTextToArray(ByVal myPath As String)
    Dim myRan As Variant
    Dim target As New FileSystemObject
    If Dir(myPath) = "" Then readTextToArray = False: Exit Function
    Dim intFino As Variant, aa As Variant, temp As Variant, a As Long, x As Long
    intFino = FreeFile
    Open myPath For Input As #intFino
    Do Until EOF(intFino)
        Line Input #intFino, aa
        temp = Split(aa, ",")
        a = UBound(temp)
        If IsEmpty(myRan) Then
            ReDim myRan(a, 0)
        End If
        ReDim Preserve myRan(a, UBound(myRan, 2) + 1)
        For x = LBound(temp) To UBound(temp)
            myRan(x, UBound(myRan, 2)) = temp(x)
        Next x
    Loop
    Close #intFino
    readTextToArray = myRan
End Function

Sub temp_sortTest()
    Dim aaaary As Variant
    aaaary = readSheetToRan3(Sheets("�[���ꗗ"), "�[�����i��", "82141V1210     ,�[�����i��,�[����,���^�p�x,���^����", "", 1)
    Stop
    Sort_Array_select aaaary, "82141V1210     ,�[����,�[�����i��", "true,true,true"
    Stop
    aaaary = array_add_column(aaaary, "width,height")
    Stop
    export_ArrayToSheet_v2 aaaary, "aaaary", True
    Stop
End Sub
'������Array�ɗ��ǉ�
Public Function array_add_column(ByVal ary As Variant, ByVal addFieldNames As String) As Variant
    Dim newAry As Variant, sp As Variant
    Dim y As Long, x As Long
    sp = Split(addFieldNames, ",")
    ReDim newAry(UBound(ary) + UBound(sp) + 1, UBound(ary, 2))
    '�ǉ������t�B�[���h����n��
    For x = LBound(sp) To UBound(sp)
        newAry(UBound(ary) + x + 1, 0) = sp(x)
    Next x
    '�l��n��
    For y = LBound(ary, 2) To UBound(ary, 2)
        For x = LBound(ary) To UBound(ary)
            newAry(x, y) = ary(x, y)
        Next x
    Next y
    array_add_column = newAry
End Function





