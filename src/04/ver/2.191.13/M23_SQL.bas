Attribute VB_Name = "M23_SQL"

Sub SQL_�z���[���擾(�z���[��RAN, ���i�i��str, �T�ustr)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim �z���[��RAN(1, 0)
    
    Dim mysql(1) As String, ����(1) As String
    '�n�_���̉�H
    mysql(0) = " SELECT [�n�_���[�����ʎq], [�n�_���n��],[�F��_],[�\��_],[�T��_],[RLTFtoPVSW_],[�n�_���}_]" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] = '" & �T�ustr & "'" & _
          " AND [�n�_���[�����ʎq] IS NOT NULL"  ' & _
          " GROUP BY  [�n�_���[�����ʎq],[�n�_���n��],[�F��_],[�n�_���}_]"
          
    '�I�_���̉�H
    mysql(1) = " SELECT �I�_���[�����ʎq, �I�_���n��,�F��_,�\��_,�T��_,RLTFtoPVSW_,�I�_���}_" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] = '" & �T�ustr & "'" & _
          " AND �I�_���[�����ʎq IS NOT NULL " '& _
          " GROUP BY  �I�_���[�����ʎq,�I�_���n��,�F��_,�I�_���}_"
    
    For a = 0 To 1
        'SQL���J��=�����ŃG���[�ɂȂ鎞�A����������PVSW_RLTF�őS���̃Z���G���^�[���s���Ȃ����񂩂�
        rs.Open mysql(a), cn, adOpenStatic
        Debug.Print rs.RecordCount
        '�z��Ɋi�[
        Do Until rs.EOF
            If rs(1).Value = "��" Then
                ����(0) = rs(0).Value
                ����(1) = rs(2).Value
                '����(2) = rs(3).Value
            Else
                ����(0) = rs(0).Value
                ����(1) = ""
                '����(2) = rs(3).Value
            End If
            If rs(0).Value = "" Then GoTo line10
            For p = 0 To UBound(�z���[��RAN, 2)
                If �z���[��RAN(0, p) = ����(0) Then
                    If �z���[��RAN(1, p) = ����(1) Then
                        GoTo line10
                    End If
                End If
            Next p
            '�����̂Ŋi�[
            ReDim Preserve �z���[��RAN(1, UBound(�z���[��RAN, 2) + j)
            For i = 0 To 1
                �z���[��RAN(i, UBound(�z���[��RAN, 2)) = ����(i)
            Next i
            j = 1
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Sub

Public Function SQL_�����@(�����@RAN)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim �����@RAN(0, 0)
    
    Dim mysql(0) As String, ����(1) As String
    '�n�_���̉�H
    mysql(0) = " SELECT [�����@]" & _
          " FROM �͈� " & _
          " WHERE [�����@] IS NOT NULL" & _
          " GROUP BY [�����@]"
          
    For a = 0 To 0
        'SQL���J��=�����ŃG���[�ɂȂ鎞�A����������PVSW_RLTF�őS���̃Z���G���^�[���s���Ȃ����񂩂�
        rs.Open mysql(a), cn, adOpenStatic
        Debug.Print rs.RecordCount
        j = 0
        '�z��Ɋi�[
        Do Until rs.EOF
            '�����̂Ŋi�[
            ReDim Preserve �����@RAN(0, j)
            �����@RAN(0, j) = rs(0)
            j = j + 1
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Function


Sub SQL_�z���[���擾_�[���p�[��(�z���[��RAN, �[��str)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim �z���[��RAN(0)
    
    Dim mysql(1) As String, ����(1) As String
    '�n�_���̉�H
    mysql(0) = " SELECT [�I�_���[�����ʎq]" & _
          " FROM �͈� " & _
          " WHERE [�n�_���[�����ʎq] = '" & �[��str & "'" & _
          " AND [�I�_���[�����ʎq] IS NOT NULL" & _
          " GROUP BY [�I�_���[�����ʎq]"
          
    '�I�_���̉�H
    mysql(1) = " SELECT [�n�_���[�����ʎq]" & _
          " FROM �͈� " & _
          " WHERE [�I�_���[�����ʎq] = '" & �[��str & "'" & _
          " AND [�n�_���[�����ʎq] IS NOT NULL" & _
          " GROUP BY [�n�_���[�����ʎq]"
    
    For a = 0 To 1
        'SQL���J��=�����ŃG���[�ɂȂ鎞�A����������PVSW_RLTF�őS���̃Z���G���^�[���s���Ȃ����񂩂�
        rs.Open mysql(a), cn, adOpenStatic
        Debug.Print rs.RecordCount
        '�z��Ɋi�[
        Do Until rs.EOF
            For i = LBound(�z���[��RAN) To UBound(�z���[��RAN)
                If rs(0) = �z���[��RAN(i) Then GoTo line10
            Next i
            '�����̂Ŋi�[
            ReDim Preserve �z���[��RAN(UBound(�z���[��RAN) + j)
            �z���[��RAN(UBound(�z���[��RAN)) = rs(0)
            j = 1
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Sub


Sub SQL_�z���[���擾_�[���p��H(RAN, �[��v, �[��str)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim RAN(6, 0)
    
    Dim mysql(1) As String, ����(1) As String
    '�n�_���̉�H
    mysql(0) = " SELECT [�n�_���[�����ʎq],[�I�_���[�����ʎq],[�F��_],[�\��_],[�T��_],[RLTFtoPVSW_],[�I�_���}_]" & _
          " FROM �͈� " & _
          " WHERE [�n�_���[�����ʎq] =" & "'" & �[��v & "'" & _
          " AND [�I�_���[�����ʎq] =" & "'" & �[��str & "'" '& _
          " AND [�n�_���[�����ʎq] IS NOT NULL"  ' & _
          " GROUP BY  [�n�_���[�����ʎq],[�n�_���n��],[�F��_],[�n�_���}_]"
          
    '�I�_���̉�H
    mysql(1) = " SELECT [�I�_���[�����ʎq], [�n�_���[�����ʎq],[�F��_],[�\��_],[�T��_],[RLTFtoPVSW_],[�n�_���}_]" & _
          " FROM �͈� " & _
          " WHERE [�I�_���[�����ʎq] =" & "'" & �[��v & "'" & _
          " AND [�n�_���[�����ʎq] =" & "'" & �[��str & "'" '& _
          " AND [�I�_���[�����ʎq] IS NOT NULL " '& _
          " GROUP BY  �I�_���[�����ʎq,�I�_���n��,�F��_,�I�_���}_"
    
    For a = 0 To 1
        'SQL���J��=�����ŃG���[�ɂȂ鎞�A����������PVSW_RLTF�őS���̃Z���G���^�[���s���Ȃ����񂩂�
        rs.Open mysql(a), cn, adOpenStatic
        '�z��Ɋi�[
        Do Until rs.EOF
            '�����\�����͊i�[���Ȃ�
            For i = 0 To UBound(RAN, 2)
                If RAN(3, i) = rs(3) Then GoTo line20
            Next i
            '�i�[
            ReDim Preserve RAN(6, UBound(RAN, 2) + j)
            For i = 0 To UBound(RAN, 1)
                RAN(i, UBound(RAN, 2)) = rs(i)
            Next i
            j = 1
line20:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Sub


Sub SQL_�z���[���擾2(�z���[��RAN, ���i�i��str, �T�ustr)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim �z���[��RAN(1, 0)
    
    Dim mysql(1) As String, ����(1) As String
    '�n�_���̉�H
    mysql(0) = " SELECT [�n�_���[�����ʎq], [�n�_���n��],[�F��_],[�\��_],[�T��_],[RLTFtoPVSW_],[�n�_���}_]" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] = '" & �T�ustr & "'" & _
          " AND [�n�_���[�����ʎq] IS NOT NULL"  ' & _
          " GROUP BY  [�n�_���[�����ʎq],[�n�_���n��],[�F��_],[�n�_���}_]"
          
    '�I�_���̉�H
    mysql(1) = " SELECT �I�_���[�����ʎq, �I�_���n��,�F��_,�\��_,�T��_,RLTFtoPVSW_,�I�_���}_" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] = '" & �T�ustr & "'" & _
          " AND �I�_���[�����ʎq IS NOT NULL " '& _
          " GROUP BY  �I�_���[�����ʎq,�I�_���n��,�F��_,�I�_���}_"
    
    For a = 0 To 1
        'SQL���J��=�����ŃG���[�ɂȂ鎞�A����������PVSW_RLTF�őS���̃Z���G���^�[���s���Ȃ����񂩂�
        rs.Open mysql(a), cn, adOpenStatic
        Debug.Print rs.RecordCount
        '�z��Ɋi�[
        Do Until rs.EOF
            If rs(1).Value = "��" Then
                ����(0) = rs(0).Value
                ����(1) = rs(2).Value
                '����(2) = rs(3).Value
            Else
                ����(0) = rs(0).Value
                ����(1) = ""
                '����(2) = rs(3).Value
            End If
            If rs(0).Value = "" Then GoTo line10
            For p = 0 To UBound(�z���[��RAN, 2)
                If �z���[��RAN(0, p) = ����(0) Then
                    GoTo line10
                End If
            Next p
            '�����̂Ŋi�[
            ReDim Preserve �z���[��RAN(1, UBound(�z���[��RAN, 2) + j)
            For i = 0 To 1
                �z���[��RAN(i, UBound(�z���[��RAN, 2)) = ����(i)
            Next i
            j = 1
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Sub


Public Function SQL_�z���}_�[���ꗗ(myBookName, ���type)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("�[���ꗗ")
        Dim myKey As Range: Set myKey = .Cells.Find("�[�����i��", , , 1)
        Dim firstRow As Long: firstRow = myKey.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        Set myKey = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myArea"
    End With
    
    Set rs = New ADODB.Recordset
    Dim mysql(0) As String
    
    '���̖��Ŏg�p����[���ꗗ��z��ɃZ�b�g
    ReDim �[���ꗗran(0)
    For r = LBound(���i�i��RAN, 2) To UBound(���i�i��RAN, 2)
        If ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "����"), r) = ���type Then
            ���i�i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
            mysql(0) = " SELECT [�[����]" & _
          " FROM myArea " & _
          " WHERE [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """""
            For a = LBound(mysql) To UBound(mysql)
                'SQL���J��=�����ŃG���[�ɂȂ鎞�A����������PVSW_RLTF�őS���̃Z���G���^�[���s���Ȃ����񂩂�
                rs.Open mysql(a), cn, adOpenStatic
                'Debug.Print rs.RecordCount
                '�z��Ɋi�[
                Do Until rs.EOF
                    For p = 0 To UBound(�[���ꗗran, 1)
                        If �[���ꗗran(p) = rs(0) Then
                            GoTo line10 '����̂Ŏ��̃��R�[�h
                        End If
                    Next p
                    '�����̂Ŋi�[
                    ReDim Preserve �[���ꗗran(UBound(�[���ꗗran, 1) + j)
                    �[���ꗗran(UBound(�[���ꗗran)) = rs(0)
                    j = 1
line10:
                    rs.MoveNext
                Loop
                rs.Close
            Next a
        End If
    Next r
    cn.Close
    
    '���̃V�[�g�ɒ[�������邩�m�F
    ReDim �[�������ꗗRAN(0): j = 0
    For p = 0 To UBound(�[���ꗗran)
        
        Set myfnd = ActiveSheet.Cells.Find(�[���ꗗran(p), , , 1)
        If myfnd Is Nothing Then
            ReDim Preserve �[�������ꗗRAN(UBound(�[�������ꗗRAN) + j)
            �[�������ꗗRAN(UBound(�[�������ꗗRAN)) = �[���ꗗran(p)
            j = 1
        End If
    Next p
    SQL_�z���}_�[���ꗗ = �[�������ꗗRAN
End Function


Sub SQL_�z����n���擾(�z����n��RAN, ���i�i��str, �T�ustr)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim �z����n��RAN(5, 0)
    Dim mysql(1) As String, ����(4) As String
    '�n�_���̉�H
    mysql(0) = " SELECT �F��_, �T��_,�n�_���[�����ʎq,�n�_���}_,�n�_���n��,����_" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] = '" & �T�ustr & "'" & _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " AND " & "�n�_���n�� = '��'"
    '�I�_���̉�H
    mysql(1) = " SELECT �F��_, �T��_,�I�_���[�����ʎq,�I�_���}_,�I�_���n��,����_" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] = '" & �T�ustr & "'" & _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " AND " & "�I�_���n�� = '��'"
    For a = 0 To 1
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        
        Do Until rs.EOF
            ReDim Preserve �z����n��RAN(rs.Fields.count - 1, j)
            For p = 0 To rs.Fields.count - 1
                �z����n��RAN(p, j) = rs(p)
            Next p
            j = j + 1
            rs.MoveNext
        Loop
        
        rs.Close
    Next a
    cn.Close

End Sub
Sub SQL_�z����n���_�Ŏ擾(RAN, ���i�i��str)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim RAN(2, 0)
    Dim mysql(1) As String, ����(4) As String
    '�n�_���̉�H
    mysql(0) = " SELECT �n�_���[�����ʎq,�n�_���L���r�e�B,�n�_���n��" & _
          " FROM �͈� " & _
          " WHERE  RLTFtoPVSW_='Found'" & _
          " AND " & "�n�_���n�� = '��'"
    '�I�_���̉�H
    mysql(1) = " SELECT �I�_���[�����ʎq,�I�_���L���r�e�B,�I�_���n��" & _
          " FROM �͈� " & _
          " WHERE  RLTFtoPVSW_='Found'" & _
          " AND " & "�I�_���n�� = '��'"
    For a = 0 To 1
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        
        Do Until rs.EOF
            ReDim Preserve RAN(rs.Fields.count - 1, j)
            For p = 0 To rs.Fields.count - 1
                RAN(p, j) = rs(p)
            Next p
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub


Sub SQL_�݊����Z�o(�݊���RAN, �݊��[��RAN, ���i�i��str)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
    cn.Open
    Set rs = New ADODB.Recordset
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim �݊���RAN(5, 0)
    Dim mysql(0) As String, ����(4) As String
    '�n�_���̉�H
    mysql(0) = " SELECT �n�_���[�����ʎq," & Chr(34) & "�n�_���L���r�e�B" & Chr(34) & ",�I�_���[�����ʎq," & Chr(34) & "�I�_���L���r�e�B" & Chr(34) & _
          " FROM �͈� " & _
          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null " & _
          " AND " & "RLTFtoPVSW_='Found'"
    '�I�_���̉�H
'    mySQL(1) = " SELECT �F��_, �T��_,�I�_���[�����ʎq,�I�_���}_,�I�_���n��" & _
'          " FROM �͈� " & _
'          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " = " & �T�ustr & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " AND " & "�I�_���n�� = '��'"
    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        j = 0
        Do Until rs.EOF
            ReDim Preserve �݊���RAN(5, j)
            For p = 0 To rs.Fields.count - 1
                �݊���RAN(p, j) = rs(p)
            Next p
            For i = LBound(�݊��[��RAN, 2) To UBound(�݊��[��RAN, 2) '�[���̍��W�𒲂ׂēo�^
                If �݊���RAN(0, j) = �݊��[��RAN(0, i) Then
                    �݊���RAN(4, j) = �݊��[��RAN(1, i)
                End If
                If �݊���RAN(2, j) = �݊��[��RAN(0, i) Then
                    �݊���RAN(5, j) = �݊��[��RAN(1, i)
                End If
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_�݊��[��(�݊��[��RAN, ���i�i��str, myBookName, ���type)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim �݊��[��RAN(3, 0)
    Dim mysql(1) As String, ����(4) As String
    '�n�_���̉�H
    
    mysql(0) = " SELECT �n�_���[�����ʎq , COUNT(1)" & _
          " FROM �͈� " & _
          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null and �n�_���[�����ʎq <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " GROUP BY �n�_���[�����ʎq"
    '�I�_���̉�H
    mysql(1) = " SELECT �I�_���[�����ʎq , COUNT(1)" & _
          " FROM �͈� " & _
          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null and �I�_���[�����ʎq <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " GROUP BY �I�_���[�����ʎq"
    j = 0
    For a = 0 To 1
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        Do Until rs.EOF
            For i = LBound(�݊��[��RAN, 2) To UBound(�݊��[��RAN, 2)
                If �݊��[��RAN(0, i) = rs(0) Then
                    �݊��[��RAN(2, i) = �݊��[��RAN(2, i) + rs(1) '�[�����J�E���g
                    flg = 1
                End If
            Next i
            
            If flg = 0 Then '�������͏���ǉ�
                ReDim Preserve �݊��[��RAN(3, j)
                
                �݊��[��RAN(0, j) = rs(0) '�[����
                Set myfound = Workbooks(myBookName).Sheets("���_" & ���type).Cells.Find(rs(0), , , 1)
                If myfound Is Nothing Then '�����W
                    �݊��[��RAN(1, j) = "�����W����"
                Else
                    �݊��[��RAN(1, j) = Workbooks(myBookName).Sheets("���_" & ���type).Cells.Find(rs(0), , , 1).Offset(, 1)
                End If
                �݊��[��RAN(2, j) = rs(1)
                j = j + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub
Public Function SQL_�[���ꗗ(�[���ꗗran, ���i�i��str, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file 'ThisWorkbook.path & "\" & ThisWorkbook.Name
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Workbooks(myBookName).Sheets("�[���ꗗ")
        Dim �[�����i�� As Range: Set �[�����i�� = .Cells.Find("�[�����i��", , , 1)
        Dim firstRow As Long: firstRow = �[�����i��.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �[�����i��.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�[�����i��.Row, .Columns.count).End(xlToLeft).Column
        lastCol = .Cells(firstRow, .Columns.count).End(xlToLeft).Column
        Set �[�����i�� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim �[���ꗗran(3, 0)
    Dim mysql(0) As String
    '�n�_���̉�H
    
    mysql(0) = " SELECT �[�����i�� ,�[����, [" & ���i�i��str & "],���^����" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] is not Null AND [" & ���i�i��str & "] <> """"" & _
          " ORDER BY [" & ���i�i��str & "] ASC"  '& _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " GROUP BY �n�_���[�����ʎq"
    '�I�_���̉�H
'    mySQL(1) = " SELECT �I�_���[�����ʎq , COUNT(1)" & _
'          " FROM �͈� " & _
'          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null and �I�_���[�����ʎq <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " GROUP BY �I�_���[�����ʎq"

    j = 0
    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        If rs(2).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
                                       
        Do Until rs.EOF
            ReDim Preserve �[���ꗗran(3, j)
            For i = LBound(�[���ꗗran, 1) To UBound(�[���ꗗran, 1)
                If IsNull(rs(i)) Then
                    �[���ꗗran(i, j) = ""
                Else
                    �[���ꗗran(i, j) = rs(i)
                End If
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
    SQL_�[���ꗗ = �[���ꗗ
End Function
Sub SQL_�T�u�[����(�T�u�[����RAN, ���i�i��str, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file 'ThisWorkbook.path & "\" & ThisWorkbook.Name
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Workbooks(myBookName).Sheets("���i�ʒ[���ꗗ")
        Dim �[�����i�� As Range: Set �[�����i�� = .Cells.Find("�[�����i��", , , 1)
        Dim firstRow As Long: firstRow = �[�����i��.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �[�����i��.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�[�����i��.Row, .Columns.count).End(xlToLeft).Column
        Set �[�����i�� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim �T�u�[����RAN(1, 0)
    Dim mysql(0) As String
    '�n�_���̉�H
    
    mysql(0) = " SELECT [" & ���i�i��str & "] ,COUNT(1)" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] is not Null AND [" & ���i�i��str & "] <> """"" & _
          " GROUP BY [" & ���i�i��str & "]" & _
          " ORDER BY [" & ���i�i��str & "] ASC"
    '�I�_���̉�H
'    mySQL(1) = " SELECT �I�_���[�����ʎq , COUNT(1)" & _
'          " FROM �͈� " & _
'          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null and �I�_���[�����ʎq <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " GROUP BY �I�_���[�����ʎq"

    j = 0
    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        If rs(0).Type <> 202 And rs(0).Type <> 200 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            ReDim Preserve �T�u�[����RAN(1, j)
            
            For i = LBound(�T�u�[����RAN, 1) To UBound(�T�u�[����RAN, 1)
                �T�u�[����RAN(i, j) = rs(i)
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Sub

Sub SQL_�T�u�[����_����m�F�ptemp(�T�u�[����RAN, ���i�i��str, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file 'ThisWorkbook.path & "\" & ThisWorkbook.Name
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("���i�ʒ[���ꗗ")
        Dim key As Range: Set key = .Cells.Find("�[�����i��", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    'Dim �͈� As DataTable
    
    ReDim �T�u�[����RAN(1, 0)
    Dim mysql(0) As String
    '�n�_���̉�H

    mysql(0) = " SELECT [" & ���i�i��str & "]" & ",count(1)" & _
          " FROM myTable" & _
          " WHERE [" & ���i�i��str & "] Is Not Null" & _
          " GROUP BY [" & ���i�i��str & "]" & _
          " ORDER BY [" & ���i�i��str & "] ASC"
          
    'mySQL(0) = " SELECT " & Chr(34) & ���i�i��str & Chr(34) & " ,COUNT(1)" & _
          " FROM �͈�" & _
          " WHERE �[�����i�� is not null" '& _
          " GROUP BY " & Chr(34) & ���i�i��str & Chr(34) & _
          " ORDER BY " & Chr(34) & ���i�i��str & Chr(34) & " ASC"
    '�I�_���̉�H
'    mySQL(1) = " SELECT �I�_���[�����ʎq , COUNT(1)" & _
'          " FROM �͈� " & _
'          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null and �I�_���[�����ʎq <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " GROUP BY �I�_���[�����ʎq"

    j = 0
    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        
        If rs.RecordCount = 0 Then Stop
        
        Do Until rs.EOF
            Debug.Print rs(0), rs(1)
            If Not IsNull(rs(0)) Then
                ReDim Preserve �T�u�[����RAN(1, j)
                For i = LBound(�T�u�[����RAN, 1) To UBound(�T�u�[����RAN, 1)
                    �T�u�[����RAN(i, j) = rs(i)
                    If rs(i) = "A" Then Stop
                Next i
                j = j + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Sub
Sub SQL_�T�u�m�F_�d���ꗗ_����m�F�ptemp(�d��RAN, ���i�i��str, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file 'ThisWorkbook.path & "\" & ThisWorkbook.Name
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    'Dim �͈� As DataTable
    
    ReDim �T�u�[����RAN(0, 0)
    Dim mysql(0) As String
    '�n�_���̉�H

    mysql(0) = " SELECT [" & ���i�i��str & "]" & _
          " FROM myTable" & _
          " WHERE [" & ���i�i��str & "] Is Not Null" & _
          " ORDER BY [" & ���i�i��str & "] ASC"
          
    'mySQL(0) = " SELECT " & Chr(34) & ���i�i��str & Chr(34) & " ,COUNT(1)" & _
          " FROM �͈�" & _
          " WHERE �[�����i�� is not null" '& _
          " GROUP BY " & Chr(34) & ���i�i��str & Chr(34) & _
          " ORDER BY " & Chr(34) & ���i�i��str & Chr(34) & " ASC"
    '�I�_���̉�H
'    mySQL(1) = " SELECT �I�_���[�����ʎq , COUNT(1)" & _
'          " FROM �͈� " & _
'          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null and �I�_���[�����ʎq <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " GROUP BY �I�_���[�����ʎq"

    j = 0
    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        
        If rs.RecordCount = 0 Then Stop
        
        Do Until rs.EOF
            Debug.Print rs(0)
            If Not IsNull(rs(0)) Then
                ReDim Preserve �T�u�[����RAN(0, j)
                If j = 45 Then Stop
                For i = LBound(�T�u�[����RAN, 1) To UBound(�T�u�[����RAN, 1)
                    �T�u�[����RAN(i, j) = rs(i)
                    If rs(i) = "31" Then Stop
                Next i
                j = j + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Sub

Sub SQL_�T�u�m�F_�d���ꗗ(�d��RAN, ���i�i��str, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�2"
    End With
    
    
    ReDim �d��RAN(8, 0)
    Dim mysql(0) As String
    mysql(0) = " SELECT [" & ���i�i��str & "],�d�����ʖ� , �n�_���[�����i�� ,�n�_���[�����ʎq , �I�_���[�����i�� ,�I�_���[�����ʎq ,����_,����_,JCDF_" & _
          " FROM �͈�2 " & _
          " WHERE " & "[RLTFtoPVSW_]='Found'" & _
          " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
          " ORDER BY [" & ���i�i��str & "] ASC"

    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        j = 0
        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            ReDim Preserve �d��RAN(8, j)
            For i = LBound(�d��RAN, 1) To UBound(�d��RAN, 1)
                �d��RAN(i, j) = rs(i)
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_�T�u�}_��Ƃߕ��i���X�g_���(RAN, ByVal ���i�i��str, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    ���i�i��str = ���i�i��str & String(15 - Len(���i�i��str), " ")
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF���[")
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim RAN(1, 0)
    Dim mysql(0) As String
    mysql(0) = " SELECT [�[�����ʎq],[EmptyPlug]" & _
          " FROM �͈� " & _
          " WHERE [EmptyPlug] IS NOT NULL AND [EmptyPlug] <> """""

    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        j = 0
        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            j = j + 1
            ReDim Preserve RAN(1, j)
            RAN(0, j) = rs(0)
            RAN(1, j) = rs(1)
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub


Sub SQL_���i�ʒ[���ꗗ(RAN, ���i�i��RAN, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(3, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For s = 1 To ���i�i��RANc
        �N�� = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "�N����"), s)
        '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
        If myTitle.Find(���i�i��RAN(1, s), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & ���i�i��RAN(1, s) & "],�n�_���[�����i�� ,�n�_���[�����ʎq ,'" & ���i�i��RAN(1, s) & "'" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        
            mysql(1) = " SELECT [" & ���i�i��RAN(1, s) & "],�I�_���[�����i�� ,�I�_���[�����ʎq ,'" & ���i�i��RAN(1, s) & "'" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        
        
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                           'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
            Do Until rs.EOF
                flg = False
                For r = LBound(RAN, 2) To UBound(RAN, 2)
                    If RAN(0, r) = rs(1) Then
                        If RAN(1, r) = rs(2) Then
                            If RAN(2, r) = rs(3) Then
                                flg = True
                                Exit For
                            End If
                        End If
                    End If
                Next r
                '�ǉ�
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve RAN(3, j)
                        RAN(0, j) = rs(1)
                        RAN(1, j) = rs(2)
                        RAN(2, j) = rs(3)
                        RAN(3, j) = �N��
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next s
    cn.Close

End Sub

Sub SQL_�d���ꗗ(RAN, ���i�i��RAN, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(7, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
    For s = 1 To ���i�i��RANc
        '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
        If myTitle.Find(���i�i��RAN(1, s), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 0
            mysql(0) = " SELECT [" & ���i�i��RAN(1, s) & "],�i��_,�T�C�Y_,�T��_,�F_,�F��_,SA,'" & ���i�i��RAN(1, s) & "'" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        
'            mySQL(1) = " SELECT [" & ���i�i��RAN(1, s) & "],�I�_���[�����i�� ,�I�_���[�����ʎq ,'" & ���i�i��RAN(1, s) & "'" & _
'                  " FROM �͈� " & _
'                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
'                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
'                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        
        
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                           'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
            Do Until rs.EOF
                flg = False
                '�o�^�����邩�m�F
                For r = LBound(RAN, 2) To UBound(RAN, 2)
                    If RAN(0, r) = rs(1) Then
                        If RAN(1, r) = rs(2) Then
                            If RAN(2, r) = rs(3) Then
                                If RAN(3, r) = rs(4) Then
                                    If RAN(4, r) = rs(5) Then
                                        If RAN(5, r) = rs(6) Then
                                            If RAN(6, r) = rs(7) Then
                                                flg = True
                                                RAN(7, r) = RAN(7, r) + 1
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next r
                '�ǉ�
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve RAN(7, j)
                        RAN(0, j) = rs(1)
                        RAN(1, j) = rs(2)
                        RAN(2, j) = rs(3)
                        RAN(3, j) = rs(4)
                        RAN(4, j) = rs(5)
                        RAN(5, j) = rs(6)
                        RAN(6, j) = rs(7) '���i�i��
                        RAN(7, j) = 1     '�g�p�ӏ���
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next s
    cn.Close

End Sub


Sub SQL_�R�l�N�^�ꗗ(RAN, ���i�i��RAN, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(4, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For s = 1 To ���i�i��RANc
        '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
        If myTitle.Find(���i�i��RAN(1, s), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & ���i�i��RAN(1, s) & "],�n�_���[�����i��,�n�_���[�����ʎq,TI1,'" & ���i�i��RAN(1, s) & "'" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
            mysql(1) = " SELECT [" & ���i�i��RAN(1, s) & "],�I�_���[�����i��,�I�_���[�����ʎq,TI2,'" & ���i�i��RAN(1, s) & "'" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                           'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
            Do Until rs.EOF
                flg = False
                '�o�^�����邩�m�F
                For r = LBound(RAN, 2) To UBound(RAN, 2)
                    If RAN(0, r) = rs(1) Then
                        If RAN(1, r) = rs(2) Then
                            If RAN(2, r) = rs(3) Then
                                If RAN(3, r) = rs(4) Then
                                    flg = True
                                    RAN(4, r) = RAN(4, r) + 1
                                End If
                            End If
                        End If
                    End If
                Next r
                '�ǉ�
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve RAN(4, j)
                        RAN(0, j) = rs(1)
                        RAN(1, j) = rs(2)
                        RAN(2, j) = rs(3)
                        RAN(3, j) = rs(4)  '���i�i��
                        RAN(4, j) = 1      '�g�p�ӏ���
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next s
    cn.Close

End Sub

Sub SQL_�}���K�C�h�o�^�ꗗ(RAN, ���i�i��RAN, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(5, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For s = 1 To ���i�i��RANc
        '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
        If myTitle.Find(���i�i��RAN(1, s), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & ���i�i��RAN(1, s) & "],�n�_���[�����i��,�n�_���[�����ʎq,TI1,'" & ���i�i��RAN(1, s) & "',TI_�n�_���}���K�C�h" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
            mysql(1) = " SELECT [" & ���i�i��RAN(1, s) & "],�I�_���[�����i��,�I�_���[�����ʎq,TI2,'" & ���i�i��RAN(1, s) & "',TI_�I�_���}���K�C�h" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                           'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
            Do Until rs.EOF
                flg = False
                '�o�^�����邩�m�F
                For r = LBound(RAN, 2) To UBound(RAN, 2)
                    If RAN(0, r) = rs(1) Then
                        If RAN(1, r) = rs(2) Then
                            If RAN(2, r) = rs(3) Then
                                If RAN(3, r) = rs(4) Then
                                    If RAN(5, r) = rs(5) Then
                                        flg = True
                                        RAN(4, r) = RAN(4, r) + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next r
                '�ǉ�
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve RAN(5, j)
                        RAN(0, j) = rs(1)
                        RAN(1, j) = rs(2)
                        RAN(2, j) = rs(3)
                        RAN(3, j) = rs(4)  '���i�i��
                        RAN(4, j) = 1      '�g�p�ӏ���
                        RAN(5, j) = rs(5)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next s
    cn.Close

End Sub
Sub SQL_YcEditor_Symbol(RAN, myBook, ���i�i��str)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Properties("Jet OLEDB:Engine Type") = 35 '����Ŏw��ł��ĂȂ�,37���ƌ^����v���Ȃ��G���[
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Debug.Print "Jet OLEDB:Engine Type", cn.Properties("Jet OLEDB:Engine Type")
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�a"
    End With
    
    With myBook.Sheets("�|�C���g�ꗗ")
        Set key = .Cells.Find("�[�����i��", , , 1)
        firstRow = key.Row
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(firstRow, key.Column), .Cells(lastRow, lastCol)).Name = "�͈�b"
        Set key = Nothing
    End With
    
    ReDim RAN(3, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)

            mysql(0) = " SELECT �͈�b.[�ȈՃ|�C���g],�͈�a.[�n�_����H����],�͈�a.[�F_],�͈�a.[�F��_]" & _
                  " FROM �͈�a INNER JOIN �͈�b" & _
                  " ON �͈�a.[�n�_���[�����ʎq] = �͈�b.[�[����] And �͈�a.[�n�_���[�����i��] = �͈�b.[�[�����i��] AND �͈�a.[�n�_���L���r�e�B] = �͈�b.[Cav] " & _
                  " WHERE " & "�͈�a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND �͈�a.[" & ���i�i��str & "] IS NOT NULL AND �͈�a.[" & ���i�i��str & "] <> """""
        
            mysql(0) = " SELECT �͈�b.�ȈՃ|�C���g,�͈�a.�n�_����H����,�͈�a.�F_,�͈�a.�F��_" & _
                  " FROM �͈�a INNER JOIN �͈�b" & _
                  " ON �͈�a.�n�_���[�����ʎq = �͈�b.�[���� And �͈�a.�n�_���[�����i�� = �͈�b.�[�����i�� AND �͈�a.�n�_���L���r�e�B = �͈�b.Cav " & _
                  " WHERE " & "�͈�a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND �͈�a.[" & ���i�i��str & "] IS NOT NULL AND �͈�a.[" & ���i�i��str & "] <> """""

                  
            mysql(1) = " SELECT �͈�b.�ȈՃ|�C���g,�͈�a.�I�_����H����,�͈�a.�F_,�͈�a.�F��_" & _
                  " FROM �͈�a INNER JOIN �͈�b" & _
                  " ON �͈�a.�I�_���[�����ʎq = �͈�b.�[���� And �͈�a.�I�_���[�����i�� = �͈�b.�[�����i�� AND �͈�a.�I�_���L���r�e�B = �͈�b.Cav " & _
                  " WHERE " & "�͈�a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND �͈�a.[" & ���i�i��str & "] IS NOT NULL AND �͈�a.[" & ���i�i��str & "] <> """""
        For k = 0 To 1
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            
            If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                           'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
            Do Until rs.EOF
                flg = False
                '�o�^�����邩�m�F
'                For r = LBound(RAN, 2) To UBound(RAN, 2)
'                    If RAN(0, r) = rs(0) Then
'                        If RAN(1, r) = rs(1) Then
'                            If RAN(2, r) = rs(2) Then
'                                If RAN(3, r) = rs(3) Then
'                                    flg = True
'                                End If
'                            End If
'                        End If
'                    End If
'                Next r
                '�ǉ�
                If flg = False Then
                    If rs(0) <> "" Then
                        j = j + 1
                        ReDim Preserve RAN(3, j)
                        RAN(0, j) = rs(0)
                        RAN(1, j) = rs(1)
                        RAN(2, j) = rs(2)
                        RAN(3, j) = rs(3)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
    cn.Close

End Sub

Sub SQL_YcEditor_WH(RAN, myBook, ���i�i��str)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim RAN(4, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
        '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
        For k = 0 To 0
            mysql(0) = " SELECT �\��_,�n�_����H����,�I�_����H����,�F_,�F��_" & _
                  " FROM �͈�" & _
                  " WHERE " & "[RLTFtoPVSW_] = 'Found'" & _
                  " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """""
                  
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                           'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
            Do Until rs.EOF
                flg = False
                '�o�^�����邩�m�F
                For r = LBound(RAN, 2) To UBound(RAN, 2)
                    If RAN(0, r) = rs(0) Then
                        If RAN(1, r) = rs(1) Then
                            If RAN(2, r) = rs(2) Then
                                If RAN(3, r) = rs(3) Then
                                    If RAN(4, r) = rs(4) Then
                                        flg = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next r
                '�ǉ�
                If flg = False Then
                    If rs(0) <> "" Then
                        j = j + 1
                        ReDim Preserve RAN(4, j)
                        RAN(0, j) = rs(0)
                        RAN(1, j) = rs(1)
                        RAN(2, j) = rs(2)
                        RAN(3, j) = rs(3)
                        RAN(4, j) = rs(4)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
    cn.Close

End Sub



Sub SQL_�}���K�C�h�ꗗ(RAN, ���i�i��RAN, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(5, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For s = 1 To ���i�i��RANc
        '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
        If myTitle.Find(���i�i��RAN(1, s), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & ���i�i��RAN(1, s) & "],�n�_���[�����i��,�n�_���[�����ʎq,TI1,'" & ���i�i��RAN(1, s) & "',TI_�n�_���}���K�C�h" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
            mysql(1) = " SELECT [" & ���i�i��RAN(1, s) & "],�I�_���[�����i��,�I�_���[�����ʎq,TI2,'" & ���i�i��RAN(1, s) & "',TI_�I�_���}���K�C�h" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                           'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
            Do Until rs.EOF
                flg = False
                '�o�^�����邩�m�F
                For r = LBound(RAN, 2) To UBound(RAN, 2)
'                    If ran(0, r) = rs(1) Then
'                        If ran(1, r) = rs(2) Then
                            If RAN(2, r) = rs(3) Then
                                If RAN(3, r) = rs(4) Then
                                    If RAN(5, r) = rs(5) Then
                                        flg = True
                                        RAN(4, r) = RAN(4, r) + 1
                                    End If
                                End If
                            End If
'                        End If
'                    End If
                Next r
                '�ǉ�
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve RAN(5, j)
                        RAN(0, j) = rs(1)
                        RAN(1, j) = rs(2)
                        RAN(2, j) = rs(3)
                        RAN(3, j) = rs(4)  '���i�i��
                        RAN(4, j) = 1      '�g�p�ӏ���
                        RAN(5, j) = rs(5)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next s
    cn.Close

End Sub

Sub SQL_�[�q�ꗗ(RAN, ���i�i��RAN, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(5, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For s = 1 To ���i�i��RANc
        '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
        If myTitle.Find(���i�i��RAN(1, s), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & ���i�i��RAN(1, s) & "],�n�_���[�q_,�n�_�����i_,�n�_����_,SM1,'" & ���i�i��RAN(1, s) & "'" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        
             mysql(1) = " SELECT [" & ���i�i��RAN(1, s) & "],�I�_���[�q_,�I�_�����i_,�I�_����_,SM2,'" & ���i�i��RAN(1, s) & "'" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
                  " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                           'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
            Do Until rs.EOF
                flg = False
                '�o�^�����邩�m�F
                For r = LBound(RAN, 2) To UBound(RAN, 2)
                    If RAN(0, r) = rs(1) Then
                        If RAN(1, r) = rs(2) Then
                            If RAN(2, r) = rs(3) Then
                                If RAN(3, r) = rs(4) Then
                                    If RAN(4, r) = rs(5) Then
                                        flg = True
                                        RAN(5, r) = RAN(5, r) + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next r
                '�ǉ�
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve RAN(5, j)
                        RAN(0, j) = rs(1)
                        RAN(1, j) = rs(2)
                        RAN(2, j) = rs(3)
                        RAN(3, j) = rs(4)
                        RAN(4, j) = rs(5) '���i�i��
                        RAN(5, j) = 1     '�g�p�ӏ���
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next s
    cn.Close

End Sub
Sub SQL_�[���T�u�ꗗ(RAN, ���i�i��str, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(2, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
    If myTitle.Find(���i�i��str, , , 1) Is Nothing Then GoTo nexts
    For k = 0 To 1
        mysql(0) = " SELECT [" & ���i�i��str & "],�n�_���[�����ʎq,�n�_���[�����i��" & _
              " FROM �͈� " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
              " ORDER BY [" & ���i�i��str & "] ASC"
    
        mysql(1) = " SELECT [" & ���i�i��str & "],�I�_���[�����ʎq,�I�_���[�����i��" & _
              " FROM �͈� " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
              " ORDER BY [" & ���i�i��str & "] ASC"
    
        'SQL���J��
        rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            '�o�^�����邩�m�F
            If Mid(rs(2), 1, 4) <> "7009" Then GoTo line10
            '�ǉ�
                If rs(1) & rs(2) <> "" Then
                    j = j + 1
                    ReDim Preserve RAN(2, j)
                    RAN(0, j) = rs(1)
                    RAN(1, j) = rs(2)
                    RAN(2, j) = rs(0)
            End If
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next k
nexts:
    cn.Close
End Sub
Sub SQL_���i�ʒ[���ꗗ_�h��(RAN, ���i�i��RAN, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(2, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For s = 1 To ���i�i��RANc
        '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
        If myTitle.Find(���i�i��RAN(1, s), , , 1) Is Nothing Then GoTo nexts
        mysql(0) = " SELECT [" & ���i�i��RAN(1, s) & "],[�n�_���[�����i��],[�n�_���[�����ʎq] " & _
              " FROM �͈� " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
              " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
    
        mysql(1) = " SELECT [" & ���i�i��RAN(1, s) & "],[�I�_���[�����i��],[�I�_���[�����ʎq] " & _
              " FROM �͈� " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & ���i�i��RAN(1, s) & "] IS NOT NULL AND [" & ���i�i��RAN(1, s) & "] <> """"" & _
              " ORDER BY [" & ���i�i��RAN(1, s) & "] ASC"
        For k = 0 To 1
        
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                           'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
            Do Until rs.EOF
                flg = False
                For r = LBound(RAN, 2) To UBound(RAN, 2)
                    If RAN(0, r) = rs(1) And RAN(1, r) = rs(2) Then
                        flg = True
                        Exit For
                    End If
                Next r
                '�ǉ�
                If flg = False Then
                    If rs(1) <> "" Then
                        j = j + 1
                        ReDim Preserve RAN(2, j)
                        RAN(0, j) = rs(1)
                        RAN(1, j) = rs(2)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next s
    cn.Close

End Sub


Sub SQL_csv�C���|�[�g(�Ώۃt�@�C��, myBookpath)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Text;HRD=YES;FMT=Delimited"
    cn.Open Left(myBookpath, InStrRev(myBookpath, "\")) & "000_�V�X�e���p�[�c\"
    Set rs = New ADODB.Recordset
    
    ReDim �d��RAN(5, 0)
    Dim mysql(0) As String
    mysql(0) = " SELECT * " & _
          " FROM " & �Ώۃt�@�C�� '& _
          " WHERE " & "[���]='�ʐ^'" ' & _
          " AND [" & ���i�i��str & "] IS NOT NULL" & _
          " ORDER BY [" & ���i�i��str & "] ASC"

    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        '���[�N�V�[�g�̒ǉ�
        For Each ws(0) In Worksheets
            If ws(0).Name = �Ώۃt�@�C�� Then
                Application.DisplayAlerts = False
                ws(0).Delete
                Application.DisplayAlerts = True
            End If
        Next ws
        Set newSheet = Worksheets.add
        newSheet.Name = �Ώۃt�@�C��
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
'                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
'        Do Until rs.EOF
'            ReDim Preserve �d��RAN(5, J)
'            For i = LBound(�d��RAN, 1) To UBound(�d��RAN, 1)
'                �d��RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop
        With newSheet
            .Cells.NumberFormat = "@"
            For i = 0 To rs.Fields.count - 1
                .Cells(1, i + 1) = rs(i).Name
            Next i
            .Range("a2").CopyFromRecordset rs
        End With
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_���i�ʒ[���ꗗ_CAV���W(RAN, ���i�i��str, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Text;HRD=YES:FMT=Delimited"
    cn.Open �A�h���X(1) & "\"
    Set rs = New ADODB.Recordset
    
    ReDim RAN(5, 0): j = 0
    Dim mysql(1) As String
    mysql(0) = " SELECT [PartName],[Cav],[Width],[Height],[EmptyPlug],[PlugColor] " & _
          " FROM CAV���W.txt" & _
          " WHERE [PartName]='" & ���i�i��str & "'" & _
             "AND [���]='�ʐ^'" & _
          " ORDER BY [Cav] ASC" ' & _
          " GROUP BY [Cav]"
    
    mysql(1) = " SELECT [PartName],[Cav],[Width],[Height],[EmptyPlug],[PlugColor] " & _
          " FROM CAV���W.txt" & _
          " WHERE [PartName]='" & ���i�i��str & "'" & _
             "AND [���]='���}'" & _
          " ORDER BY [CAV] ASC" '& _
          " GROUP BY [PartName],[Cav]"
    For a = 0 To 1
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
        If rs.RecordCount = 0 Then GoTo line20
        Do Until rs.EOF
            j = j + 1
            ReDim Preserve RAN(5, j)
            RAN(0, j) = rs(0)
            RAN(1, j) = rs(1)
            RAN(2, j) = rs(2)
            RAN(3, j) = rs(3)
            RAN(4, j) = rs(4)
            RAN(5, j) = rs(5)
            '.Cells(1, i + 1) = rs(i).Name
            '.Range("a2").CopyFromRecordset rs
            rs.MoveNext
        Loop
line20:
        rs.Close
        If j > 0 Then GoTo line40
    Next a
line40:
    cn.Close

End Sub

Public Function SQL_���i�ʒ[���ꗗ_CAV���W2(RAN, ���i�i��str, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Text;HRD=YES:FMT=Delimited"
    
    cn.Open �A�h���X(1) & "\200_CAV���W\"
    Set rs = New ADODB.Recordset
    
    ReDim RAN(5, 0): j = 0
    Dim mysql(1) As String
    mysql(0) = " SELECT [PartName],[Cav],[Width],[Height],[EmptyPlug],[PlugColor] " & _
          " FROM " & "'" & ���i�i��str & "'" '& _
          " WHERE [PartName]='" & ���i�i��str & "'" & _
          " ORDER BY [Cav] ASC" ' & _
          " GROUP BY [Cav]"
    
    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
        If rs.RecordCount = 0 Then GoTo line20
        Do Until rs.EOF
            j = j + 1
            ReDim Preserve RAN(5, j)
            RAN(0, j) = rs(0)
            RAN(1, j) = rs(1)
            RAN(2, j) = rs(2)
            RAN(3, j) = rs(3)
            RAN(4, j) = rs(4)
            RAN(5, j) = rs(5)
            '.Cells(1, i + 1) = rs(i).Name
            '.Range("a2").CopyFromRecordset rs
            rs.MoveNext
        Loop
line20:
        rs.Close
        If j > 0 Then GoTo line40
    Next a
line40:
    cn.Close
    SQL_���i�ʒ[���ꗗ_CAV���W2 = j
End Function

Sub SQL_�T�u�i���o�[���_�f�[�^�쐬(���i�i��RAN, mySheet, temp�A�h���X, ByVal myPosSP As Variant, ByVal kumitateList As Variant)
    
     '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With mySheet
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim RAN(2, 0)
    j = 0
    Dim mysql(0) As String
    
    For r = LBound(���i�i��RAN, 2) + 1 To UBound(���i�i��RAN, 2)
        �T�u��� = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "�T�u"), r)
        If �T�u��� = "1" Then
            �i��str = ���i�i��RAN(���i�i��RAN_read(���i�i��RAN, "���C���i��"), r)
    
            mysql(0) = " SELECT [" & �i��str & "],left(�d�����ʖ�,4),'" & Replace(�i��str, " ", "") & "'" & _
                  " FROM �͈� " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & �i��str & "] IS NOT NULL AND [" & �i��str & "] <> """"" & _
                  " ORDER BY [�d�����ʖ�] ASC"
            For a = 0 To 0
                'SQL���J��
                rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
                If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                               'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
                Do Until rs.EOF
                    ReDim Preserve RAN(2, j)
                    For i = LBound(RAN, 1) To UBound(RAN, 1)
                        RAN(i, j) = rs(i)
                    Next i
                    j = j + 1
                    rs.MoveNext
                Loop
                rs.Close
            Next a
        End If
    Next r
    cn.Close
    
    '�e�L�X�g�t�@�C���ɂ��ďo��
    Dim lntFlNo As Integer: lntFlNo = FreeFile
    Dim outPutAddress As String: outPutAddress = temp�A�h���X
    Open outPutAddress For Output As #lntFlNo
    
    ���� = Now: Dim myLine
    For i = LBound(RAN, 2) To UBound(RAN, 2)
        �T�u�l = RAN(0, i)
        �\�� = RAN(1, i)
        ���i = RAN(2, i)
        For r = LBound(kumitateList, 2) + 1 To UBound(kumitateList, 2)
            myLine = Empty
            For ii = LBound(myPosSP) To UBound(myPosSP)
                If myPosSP(ii) <> "" Then
                    Select Case ii
                        Case 1
                            myVal = ���i
                        Case 3
                            myVal = �\��
                        Case 4
                            myVal = �T�u�l
                        Case 5
                            myVal = kumitateList(0, r)
                        Case Else
                            myVal = "" '�ؒf�Ɛݕ�_0��2
                    End Select
                    myLine = myLine & Chr(34) & myVal & Chr(34) & Chr(44)
                End If
            Next ii
            myLine = myLine & ����
            Print #lntFlNo, myLine
        Next r
    Next i
    
    Close lntFlNo

End Sub


Sub SQL_�ύX�˗�_����(���i�i��RAN, �����ύXRAN, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    Dim mysql(0) As String
    
    mysql(0) = "SELECT "
    For i = 1 To ���i�i��RANc
        mysql(0) = mysql(0) & "[" & ���i�i��RAN(1, i - 1) & "],"
    Next i
    
    ReDim �����ύXRAN(���i�i��RANc + 6, 0)
    mysql(0) = mysql(0) & "�\��_,�n�_����H����, �I�_����H����, ����_ ,������_ ,RLTFtoPVSW_,���l_" & _
          " FROM myTable " & _
          " WHERE " & "[RLTFtoPVSW_]='Found'" & _
          " AND [������_] IS NOT NULL"

    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        j = 0
        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            ReDim Preserve �����ύXRAN(���i�i��RANc + 6, j)
            For i = LBound(�����ύXRAN, 1) To UBound(�����ύXRAN, 1)
                �����ύXRAN(i, j) = rs(i)
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_�[���ꗗ_2(���i�i��RAN, �d���ꗗRAN, �[���ꗗran, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF_temp")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    Dim mysql(0) As String
    
    mysql(0) = "SELECT "
    For i = 1 To ���i�i��RANc
        mysql(0) = mysql(0) & "[" & ���i�i��RAN(1, i - 1) & "],"
    Next i
    
    ReDim �d���ꗗRAN(���i�i��RANc + 9, 0)
    ReDim �[���ꗗran(0)
    mysql(0) = mysql(0) & "�\��_,�n�_����H����, �I�_����H����, �n�_���[�����ʎq, �I�_���[�����ʎq,�n�_���L���r�e�B,�I�_���L���r�e�B,����_,������_ ,RLTFtoPVSW_,���l_" & _
          " FROM myTable " & _
          " WHERE " & "[RLTFtoPVSW_]='Found'" '& _
          " AND [������_] IS NOT NULL"

    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim j As Long: j = 0
        Dim jj As Long: jj = 0
        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            '���type�̑Ώۂɂ��邩�m�F
            findFlg = False
            For i = 1 To ���i�i��RANc
                If Not IsNull(rs(i - 1)) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            
            If findFlg = False Then
                GoTo line20
            End If
            
            '�ǉ�
            ReDim Preserve �d���ꗗRAN(���i�i��RANc + 9, j)
            
            For i = 1 To ���i�i��RANc
                �d���ꗗRAN(i - 1, j + 0) = rs(i - 1)
            Next i
                '�n�_
                �d���ꗗRAN(���i�i��RANc + 0, j + 0) = rs(���i�i��RANc + 0) '�\��
                �d���ꗗRAN(���i�i��RANc + 1, j + 0) = rs(���i�i��RANc + 1) '��
                �d���ꗗRAN(���i�i��RANc + 2, j + 0) = rs(���i�i��RANc + 2)
                �d���ꗗRAN(���i�i��RANc + 3, j + 0) = rs(���i�i��RANc + 3) '�[��
                �d���ꗗRAN(���i�i��RANc + 4, j + 0) = rs(���i�i��RANc + 4)
                �d���ꗗRAN(���i�i��RANc + 5, j + 0) = rs(���i�i��RANc + 5) 'cav
                �d���ꗗRAN(���i�i��RANc + 6, j + 0) = rs(���i�i��RANc + 6)
                �d���ꗗRAN(���i�i��RANc + 7, j + 0) = rs(���i�i��RANc + 7) '����_
                �d���ꗗRAN(���i�i��RANc + 8, j + 0) = rs(���i�i��RANc + 8) '������_
                �d���ꗗRAN(���i�i��RANc + 9, j + 0) = rs(���i�i��RANc + 10) '���l_
                
            '�n�_�[���������ǉ�
            For i = LBound(�[���ꗗran) To UBound(�[���ꗗran)
                findFlg = False
                If �[���ꗗran(i) = rs(���i�i��RANc + 3) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            If findFlg = False Then
                ReDim Preserve �[���ꗗran(jj)
                �[���ꗗran(jj) = rs(���i�i��RANc + 3)
                jj = jj + 1
            End If
            '�I�_�[���������ǉ�
            For i = LBound(�[���ꗗran) To UBound(�[���ꗗran)
                findFlg = False
                If �[���ꗗran(i) = rs(���i�i��RANc + 4) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            If findFlg = False Then
                ReDim Preserve �[���ꗗran(jj)
                �[���ꗗran(jj) = rs(���i�i��RANc + 4)
                jj = jj + 1
            End If
            j = j + 1
line20:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub
Sub SQL_�n���}�쐬_1(���i�i��RAN, �n���}�쐬RAN, �[��, myBook, newSheet)
    
    Call SQL_csv�C���|�[�g("���ޏڍ�.txt", myBook.Path)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open ThisWorkbook.FullName
    Set rs = New ADODB.Recordset
    
    'myTable0
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable0"
    End With
    
    'myTable1
    With myBook.Sheets("�|�C���g�ꗗ")
        Set key = .Cells.Find("�[�����i��", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable1"
    End With
    
    'myTable2
    With myBook.Sheets("���i�ʒ[���ꗗ")
        Set key = .Cells.Find("�h���R�l�N�^�i��", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(firstRow, key.Column), .Cells(lastRow, lastCol)).Name = "myTable2"
        Set key = Nothing
    End With
    
    'myTable3
    With myBook.Sheets("���ޏڍ�.txt")
        Set key = .Cells.Find("���i�i��_", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(firstRow, key.Column), .Cells(lastRow, lastCol)).Name = "myTable3"
        Set key = Nothing
    End With
    
    Dim �d���ꗗRAN() As String
    ReDim �d���ꗗRAN(���i�i��RANc + 11, 0)
    Dim mysql(1) As String
    
    For a = LBound(mysql) To UBound(mysql)
        mysql(a) = "SELECT "
        For i = 1 To ���i�i��RANc
            mysql(a) = mysql(a) & "[" & ���i�i��RAN(1, i - 1) & "],"
        Next i
    Next a
    
    mysql(0) = mysql(0) & "�\��_,�n�_����H����, �n�_���[�����ʎq,�n�_���L���r�e�B,�n�_���[�����i��,����_,������_ ,RLTFtoPVSW_,�n�_���}_,�F��_,�i��_,�T��_,���[�n��,���[���[�q,�n�_���n��,�n�_������_,�n�_����_,'�n' AS ��" & _
                          ",b.[�|�C���g1]" & _
                          ",c.[EmptyPlug],c.[PlugColor]" & _
                          ",d.[�R�l�N�^�ɐ�_]" & _
          " FROM (((myTable0 AS a" & _
          " LEFT OUTER JOIN myTable1 AS b " & _
          " ON a.[�n�_���[�����i��] = b.[�[�����i��] AND a.[�n�_���[�����ʎq] = b.[�[����] AND a.[�n�_���L���r�e�B] = b.[Cav])" & _
          " LEFT OUTER JOIN myTable2 AS c " & _
          " ON a.[�n�_���[�����i��] = c.[�h���R�l�N�^�i��] AND a.[�n�_���[�����ʎq] = c.[�[����_] AND a.[�n�_���L���r�e�B] = c.[Cav])" & _
          " LEFT OUTER JOIN myTable3 AS d " & _
          " ON a.[�n�_���[�����i��] = d.[���i�i��_] )" & _
          " WHERE " & "a.[RLTFtoPVSW_]='Found' AND a.[�n�_���[�����ʎq] is not Null AND a.[�n�_���L���r�e�B] is not Null"

    mysql(1) = mysql(1) & "�\��_,�I�_����H����, �I�_���[�����ʎq,�I�_���L���r�e�B,�I�_���[�����i��,����_,������_ ,RLTFtoPVSW_,�I�_���}_,�F��_,�i��_,�T��_,���[�n��,���[���[�q,�I�_���n��,�I�_������_,�I�_����_,'�I' AS ��" & _
                          ",b.[�|�C���g1]" & _
                          ",c.[EmptyPlug],c.[PlugColor]" & _
                          ",d.[�R�l�N�^�ɐ�_]" & _
          " FROM (((myTable0 AS a" & _
          " LEFT OUTER JOIN myTable1 AS b " & _
          " ON a.[�I�_���[�����i��] = b.[�[�����i��] AND a.[�I�_���[�����ʎq] = b.[�[����] AND a.[�I�_���L���r�e�B] = b.[Cav])" & _
          " LEFT OUTER JOIN myTable2 AS c " & _
          " ON a.[�I�_���[�����i��] = c.[�h���R�l�N�^�i��] AND a.[�I�_���[�����ʎq] = c.[�[����_] AND a.[�I�_���L���r�e�B] = c.[Cav])" & _
          " LEFT OUTER JOIN myTable3 AS d " & _
          " ON a.[�I�_���[�����i��] = d.[���i�i��_] )" & _
          " WHERE " & "a.[RLTFtoPVSW_]='Found' AND a.[�I�_���[�����ʎq] is not Null AND a.[�I�_���L���r�e�B] is not Null"

    For a = LBound(mysql) To UBound(mysql)
        For i = 1 To ���i�i��RANc
            If i = 1 Then
                mysql(a) = mysql(a) & " AND [" & ���i�i��RAN(1, i - 1) & "] is not null"
            Else
                mysql(a) = mysql(a) & " OR [" & ���i�i��RAN(1, i - 1) & "] is not null"
            End If
        Next i
    Next a
          
    'mySQL(0) = mySQL(0) & " ORDER BY [�n�_���[�����ʎq] ASC , [�n�_���L���r�e�B] ASC"

    For a = LBound(mysql) To UBound(mysql)
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim j As Long: j = 0
        Dim jj As Long: jj = 0
        '���[�N�V�[�g�̒ǉ�
        If a = LBound(mysql) Then
            For Each ws(0) In Worksheets
                If ws(0).Name = "�n���}temp" Then
                    Application.DisplayAlerts = False
                    ws(0).Delete
                    Application.DisplayAlerts = True
                End If
            Next ws
            Set newSheet = Worksheets.add
            newSheet.Name = "�n���}temp"
        End If
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
'                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
'        Do Until rs.EOF
'            ReDim Preserve �d��RAN(5, J)
'            For i = LBound(�d��RAN, 1) To UBound(�d��RAN, 1)
'                �d��RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop
        With newSheet
            .Cells.NumberFormat = "@"
            For i = 0 To rs.Fields.count - 1
                .Cells(1, i + 1) = Replace(Replace(rs(i).Name, "�n�_��", ""), "�I�_��", "")
            Next i
            lastRow = .Cells(.Rows.count, .Cells.Find("�\��_", , , 1).Column).End(xlUp).Row + 1
            .Cells(lastRow, 1).CopyFromRecordset rs
        End With
        Debug.Print rs.RecordCount
        rs.Close
    Next a
    cn.Close
    
End Sub

Sub SQL_�n���}�쐬_2(���i�i��RAN, myBook, newSheet)
    
    Call SQL_csv�C���|�[�g("CAV���W.txt", myBook.Path)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open ThisWorkbook.FullName
    Set rs = New ADODB.Recordset
    
    'myTable0
    With newSheet
        Dim firstRow As Long: firstRow = 1
        Dim lastRow0 As Long: lastRow0 = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(firstRow, 1), .Cells(lastRow0, lastCol)).Name = "myTable0"
    End With
    
    'myTable1
    With myBook.Sheets("CAV���W.txt")
        Set key = .Cells.Find("PartName", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable1"
    End With
    
    Dim mysql(0) As String
    mysql(0) = "SELECT a.*,b.[x],b.[���]" & _
          " FROM myTable1 AS a" & _
          " LEFT OUTER JOIN myTable0 AS b " & _
          " ON a.[�[�����i��] = b.[PartName] AND a.[�L���r�e�B] = b.[Cav] " & _
          " WHERE b.[���] = '�ʐ^'" 'a.[RLTFtoPVSW_]='Found' AND a.[�n�_���[�����ʎq] is not Null AND a.[�n�_���L���r�e�B] is not Null"
          
    'mySQL(1) = "SELECT a.* " & _
                     ",b.[x] ,b.[���]" & _
          " FROM myTable0 AS a" & _
          " LEFT OUTER JOIN myTable1 AS b " & _
          " ON a.[�[�����i��] = b.[PartName] AND a.[�L���r�e�B] = b.[Cav]" & _
          " WHERE b.[���] = '���}'" 'a.[RLTFtoPVSW_]='Found' AND a.[�n�_���[�����ʎq] is not Null AND a.[�n�_���L���r�e�B] is not Null"
          
    'mySQL(0) = mySQL(0) & " ORDER BY [�n�_���[�����ʎq] ASC , [�n�_���L���r�e�B] ASC"

    For a = LBound(mysql) To UBound(mysql)
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim j As Long: j = 0
        Dim jj As Long: jj = 0
        '�Z���̒l���폜
        If a = LBound(mysql) Then
            '���[�N�V�[�g�̒ǉ�
            If a = LBound(mysql) Then
                For Each ws(0) In Worksheets
                    If ws(0).Name = "�n���}temp1" Then
                        Application.DisplayAlerts = False
                        ws(0).Delete
                        Application.DisplayAlerts = True
                    End If
                Next ws
                Set newSheet = Worksheets.add
                newSheet.Name = "�n���}temp1"
            End If
        End If
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
'                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
'        Do Until rs.EOF
'            ReDim Preserve �d��RAN(5, J)
'            For i = LBound(�d��RAN, 1) To UBound(�d��RAN, 1)
'                �d��RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop

'        With newSheet
'            .Cells.NumberFormat = "@"
'            For i = 0 To rs.Fields.count - 1
'                .Cells(1, i + 1) = Replace(Replace(rs(i).Name, "�n�_��", ""), "�I�_��", "")
'            Next i
'            lastRow = .UsedRange.Rows.count + 1
'            Debug.Print rs.RecordCount
'            .Cells(lastRow, 1).CopyFromRecordset rs
'        End With
        rs.Close
    Next a
    cn.Close
    
End Sub

Sub SQL_�T�u�i���o�[���_�f�[�^�X�V(temp�A�h���X, temp�A�h���X2, temp�A�h���X3, ByVal mySQLon)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    '�w�b�_�[�̖����e�L�X�g�t�@�C���̎� 12.0���ƃt�B�[���h����F1�łƂ�Ȃ�
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=NO;FMT=Delimited"
    cn.Open Left(temp�A�h���X, InStrRev(temp�A�h���X, "\") - 1)
    Set rs = New ADODB.Recordset

    Dim mysql(2) As String
    'change(1�Ɋ܂ނ���s�v�ɂȂ���)
    mysql(0) = " SELECT b.* " & _
          " FROM " & Mid(temp�A�h���X2, InStrRev(temp�A�h���X2, "\") + 1) & " as b" & _
          " INNER JOIN " & Mid(temp�A�h���X, InStrRev(temp�A�h���X, "\") + 1) & " as a" & _
          " ON a.F2 = b.F2 AND a.F4 = b.F4 "
    'new��change
    mysql(1) = " SELECT b.* " & _
          " FROM " & Mid(temp�A�h���X2, InStrRev(temp�A�h���X2, "\") + 1) & " as b" & _
          " LEFT OUTER JOIN " & Mid(temp�A�h���X, InStrRev(temp�A�h���X, "\") + 1) & " as a" & _
          mySQLon(0)
    'old
    mysql(2) = " SELECT a.* " & _
          " FROM " & Mid(temp�A�h���X2, InStrRev(temp�A�h���X2, "\") + 1) & " as b" & _
          " RIGHT OUTER JOIN " & Mid(temp�A�h���X, InStrRev(temp�A�h���X, "\") + 1) & " as a" & _
          mySQLon(1)
    
    For a = 1 To UBound(mysql)
        'SQL���J��
        'cn.Execute mySQL(0)
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        If a = 1 Then ReDim �T�u���ran(rs.Fields.count - 1, 0): j = 0
        'Sheets("Sheet1").Cells.ClearContents
        Do Until rs.EOF
            ReDim Preserve �T�u���ran(rs.Fields.count - 1, j)
            For i = 0 To rs.Fields.count - 1
                'Sheets("Sheet1").Cells(J + 1, i + 1) = rs(i).Value
                �T�u���ran(i, j) = rs(i).Value
            Next i
            j = j + 1
            'Range("a2").CopyFromRecordset rs
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
    '�t�@�C���쐬
    Dim lntFlNo As Integer: lntFlNo = FreeFile
    Open temp�A�h���X3 For Output As #lntFlNo
    
    Dim �T�u�l As String, �\�� As String, ���i As String
    Dim ���� As Date
    Dim X As Long, Y As Long, fndX As Long
    
    For X = LBound(�T�u���ran, 2) To UBound(�T�u���ran, 2)
        If Not IsNull(�T�u���ran(1, X)) Then
        myLine = Empty
        For xx = LBound(�T�u���ran) To UBound(�T�u���ran)
            If xx <> UBound(�T�u���ran) Then
                myLine = myLine & Chr(34) & �T�u���ran(xx, X) & Chr(34) & Chr(44)
            Else
                myLine = myLine & �T�u���ran(xx, X) '�Ō�͓���
            End If
        Next xx
        Print #lntFlNo, myLine
line20:
        End If
    Next X
    
    Close #lntFlNo
    
End Sub
Public Function SQL_MD�t�@�C���ǂݍ���_���(���i�i��str, �ݕ�str, myRan)
    ���i�i��str = Replace(���i�i��str, " ", "")
    temp�A�h���X1 = ThisWorkbook.Path & "\08_MD\" & ���i�i��str & "_" & �ݕ�str & "_MD" & "\004Term.csv"
    temp�A�h���X2 = ThisWorkbook.Path & "\08_MD\" & ���i�i��str & "_" & �ݕ�str & "_MD" & "\006Cone.csv"
    If Dir(temp�A�h���X1) = "" Then Exit Function
    If Dir(temp�A�h���X2) = "" Then Exit Function
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=YES;FMT=Delimited"
    cn.Open Left(temp�A�h���X1, InStrRev(temp�A�h���X1, "\") - 1)
    Set rs = New ADODB.Recordset

    Dim mysql(0) As String
    mysql(0) = " SELECT a.���i�i��,a.�T�u�ԍ�,a.�L���r�e�B�ԍ�,a.�����H��,b.�R�l�N�^�ԍ� ,b.���i�i��" & _
          " FROM " & Mid(temp�A�h���X1, InStrRev(temp�A�h���X1, "\") + 1) & " as a" & _
          " INNER JOIN " & Mid(temp�A�h���X2, InStrRev(temp�A�h���X2, "\") + 1) & " as b" & _
          " ON a.��t����h�c = b.�h�c " 'AND a.F4 = b.F4 "
    j = 0
    For a = 0 To UBound(mysql)
        'SQL���J��
        'cn.Execute mySQL(0)
        On Error Resume Next
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        If Err.Number = -2147467259 Then GoTo line20
        On Error GoTo 0
        If a = 0 Then ReDim myRan(rs.Fields.count, 0): j = 0
        
        Do Until rs.EOF
            ReDim Preserve myRan(rs.Fields.count, j)
            For i = 0 To rs.Fields.count - 1
                'Sheets("Sheet1").Cells(J + 1, i + 1) = rs(i).Value
                myRan(i, j) = rs(i).Value
            Next i
            j = j + 1
            'Range("a2").CopyFromRecordset rs
            rs.MoveNext
        Loop
        rs.Close
    Next a
line20:
    cn.Close
    
    SQL_MD�t�@�C���ǂݍ���_��� = UBound(myRan, 2)
End Function

Sub SQL_test()
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    '�w�b�_�[�̖����e�L�X�g�t�@�C���̎� 12.0���ƃt�B�[���h����F1�łƂ�Ȃ�
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=NO;FMT=Delimited"
    cn.Open "D:\04_���i�̓���\028_675W_543B"
    Set rs = New ADODB.Recordset
    
    '�������̎��A�X�V
    ReDim �d��RAN(5, 0)
    Dim mysql(0) As String
    mysql(0) = " SELECT * " & _
          " FROM efu_subNo_temp3.txt " & _
          " WHERE F6 in " & _
          " ( SELECT MAX(F6) FROM efu_subNo_temp3.txt GROUP BY F2,F4 ORDER BY F2,F4)" '& _
          " INNER JOIN " & Mid(temp�A�h���X, InStrRev(temp�A�h���X, "\") + 1) & " as b" & _
          " ON a.F2=b.F2 AND a.F4 = b.F4" '& _
          " SET a.F4 = b.F4" & _
          " WHERE a.F2=b.F2 AND a.F4 = b.F4" ' & _
          " AND [" & ���i�i��str & "] IS NOT NULL" & _
          " ORDER BY [" & ���i�i��str & "] ASC"
          'mySQL(0) = "SELECT MAX(F6),F2,F4 FROM efu_subNo_temp3.txt GROUP BY F2,F4 ORDER BY F2,F4"

    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        '���[�N�V�[�g�̒ǉ�
        For Each ws(0) In Worksheets
            If ws(0).Name = �Ώۃt�@�C�� Then
                Application.DisplayAlerts = False
                ws(0).Delete
                Application.DisplayAlerts = True
            End If
        Next ws
        Set newSheet = Worksheets.add
        newSheet.Name = �Ώۃt�@�C��
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
'                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
'        Do Until rs.EOF
'            ReDim Preserve �d��RAN(5, J)
'            For i = LBound(�d��RAN, 1) To UBound(�d��RAN, 1)
'                �d��RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop
        With newSheet
            .Cells.NumberFormat = "@"
            For i = 0 To rs.Fields.count - 1
                .Cells(1, i + 1) = rs(i).Name
            Next i
            .Range("a2").CopyFromRecordset rs
        End With
        rs.Close
    Next a
    cn.Close

End Sub

Sub Sample01forExcel()
Dim con As Object, rec As Object

    Set con = CreateObject("ADODB.Connection")
        With con
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\04_���i�̓���\028_675W_543B;" _
                                & "Extended Properties='text;HDR=No;FMT=Delimited'"
            .Open
        End With
    
    Set rec = CreateObject("ADODB.Recordset")
        rec.Open "select * from efu_subNo_temp2.txt as a where a.[F2] ='821113B300'", con
        Debug.Print rec(1) '�ŏ��̃��R�[�h��1��ڂ̒l��\��

End Sub

Sub SQL_CAV���W�擾(���i�i��RAN, myBook, newSheet)
    
    Call SQL_csv�C���|�[�g("CAV���W.txt", myBook.Path)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open ThisWorkbook.FullName
    Set rs = New ADODB.Recordset
    
    'myTable1
    With newSheet
        Dim firstRow As Long: firstRow = 1
        Dim lastRow0 As Long: lastRow0 = .UsedRange.Rows.count
        Dim lastCol0 As Long: lastCol0 = .Cells(1, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(firstRow, 1), .Cells(lastRow0, lastCol0)).Name = "myTable1"
    End With
    
    'myTable0
    With myBook.Sheets("CAV���W.txt")
        Set key = .Cells.Find("PartName", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable0"
    End With
    
    Dim mysql(0) As String
    mysql(0) = "SELECT a.[PartName],a.[Cav],a.[X],a.[Y],a.[Width],a.[Height],a.[�`��],a.[���],a.[Angle],a.[Width(mm)],a.[Category],a.[Rock]" & _
          " FROM myTable0 AS a" & _
          " LEFT JOIN myTable1 AS b " & _
          " ON a.[PartName] = b.[�[�����i��] AND a.[Cav] = b.[�L���r�e�B]" & _
          " WHERE a.[PartName] is not Null" 'a.[RLTFtoPVSW_]='Found' AND a.[�n�_���[�����ʎq] is not Null AND a.[�n�_���L���r�e�B] is not Null"
          
    'mySQL(1) = "SELECT a.* " & _
                     ",b.[x] ,b.[���]" & _
          " FROM myTable0 AS a" & _
          " LEFT OUTER JOIN myTable1 AS b " & _
          " ON a.[�[�����i��] = b.[PartName] AND a.[�L���r�e�B] = b.[Cav]" & _
          " WHERE b.[���] = '���}'" 'a.[RLTFtoPVSW_]='Found' AND a.[�n�_���[�����ʎq] is not Null AND a.[�n�_���L���r�e�B] is not Null"
          
    'mySQL(0) = mySQL(0) & " ORDER BY [�n�_���[�����ʎq] ASC , [�n�_���L���r�e�B] ASC"

    For a = LBound(mysql) To UBound(mysql)
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim addCol() As Long, �ǉ�F
        Dim cav As String
        With newSheet
            �ǉ�F = "X,Y,Width,Height,�`��,���,Angle,Width(mm),Category,Rock"
            ReDim addCol(rs.Fields.count - 1)

            For X = 1 To rs.Fields.count
                If InStr(�ǉ�F, rs(X - 1).Name) > 0 Then
                    addCol(X - 1) = .Cells(1, .Columns.count).End(xlToLeft).Column + 1
                    .Cells(1, addCol(X - 1)) = rs(X - 1).Name
                Else
                    addCol(X - 1) = 0
                End If
            Next X
            ���Col = .Rows(1).Find("�[�����i��", , , 1).Column
            cavCol = .Rows(1).Find("�L���r�e�B", , , 1).Column
            For i = 2 To lastRow
                ��� = .Cells(i, ���Col)
                cav = .Cells(i, cavCol)
                If ��� <> "" Then
                    rs.filter = "(PartName = '" & ��� & "') AND (Cav = '" & cav & "') AND (��� = '" & "�ʐ^')"
                    If rs.EOF = True Then rs.filter = "(PartName = '" & ��� & "') AND (Cav = '" & cav & "') AND (��� = '" & "���}')"
                    For X = 1 To rs.Fields.count
                        If addCol(X - 1) <> 0 Then
                            .Cells(i, addCol(X - 1)) = rs(X - 1)
                        End If
                    Next X
                End If
'                rs.Find "(PartName = '7283702640') AND (Cav = '1')", 0, adSearchForward
'                rs.Find "(PartName = '" & ��� & "') AND (Cav = '" & Cav & "')", 0, adSearchForward
'                Do Until rs.EOF
'
'                Loop
            Next i
        End With
'        Dim J As Long: J = 0
'        Dim jj As Long: jj = 0
        '�Z���̒l���폜
'        If a = LBound(mySQL) Then
'            '���[�N�V�[�g�̒ǉ�
'            If a = LBound(mySQL) Then
'                For Each ws In Worksheets
'                    If ws.Name = "�n���}temp1" Then
'                        Application.DisplayAlerts = False
'                        ws.Delete
'                        Application.DisplayAlerts = True
'                    End If
'                Next ws
'                Set newSheet = Worksheets.Add
'                newSheet.Name = "�n���}temp1"
'            End If
'        End If
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
'                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
'        Do Until rs.EOF
'            ReDim Preserve �d��RAN(5, J)
'            For i = LBound(�d��RAN, 1) To UBound(�d��RAN, 1)
'                �d��RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop

'        With newSheet
'            .Cells.NumberFormat = "@"
'            For i = 0 To rs.Fields.count - 1
'                .Cells(1, i + 1) = Replace(Replace(rs(i).Name, "�n�_��", ""), "�I�_��", "")
'            Next i
'            lastRow = .UsedRange.Rows.count + 1
'            Debug.Print rs.RecordCount
'            .Cells(lastRow, 1).CopyFromRecordset rs
'        End With
        rs.Close
    Next a
    cn.Close
    
End Sub
Sub SQL_���[�J���d���T�u�i���o�[�擾(RAN, ���i�i��)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    '�w�b�_�[�̖����e�L�X�g�t�@�C���̎� 12.0���ƃt�B�[���h����F1�łƂ�Ȃ�
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=NO;FMT=Delimited"
    cn.Open Left(�A�h���X(2), InStrRev(�A�h���X(2), "\") - 1)
    Set rs = New ADODB.Recordset

    Dim mysql(0) As String
    
    mysql(0) = " SELECT * " & _
          " FROM " & Mid(�A�h���X(2), InStrRev(�A�h���X(2), "\") + 1) & _
          " WHERE F1 = '" & ���i�i�� & "' "
          
    For a = 0 To UBound(mysql)
        'SQL���J��
        'cn.Execute mySQL(0)
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        If a = 0 Then ReDim RAN(rs.Fields.count, 0): j = 0
        
        'Sheets("Sheet1").Cells.ClearContents
        Do Until rs.EOF
            ReDim Preserve RAN(rs.Fields.count, j)
            For i = 0 To rs.Fields.count - 1
                'Sheets("Sheet1").Cells(J + 1, i + 1) = rs(i).Value
                RAN(i, j) = rs(i).Value
            Next i
            j = j + 1
            'Range("a2").CopyFromRecordset rs
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Sub

Public Function SQL_���[�J���[���T�u�i���o�[�擾(RAN, ���i�i��)
    
    If Dir(Left(�A�h���X(2), InStrRev(�A�h���X(2), "\") - 1) & "\TerminalSubNumber\" & Replace(���i�i��, " ", "") & ".txt") = "" Then
        SQL_���[�J���[���T�u�i���o�[�擾 = False
        Exit Function
    End If
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    '�w�b�_�[�̖����e�L�X�g�t�@�C���̎� 12.0���ƃt�B�[���h����F1�łƂ�Ȃ�
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=NO;FMT=Delimited"
    cn.Open Left(�A�h���X(2), InStrRev(�A�h���X(2), "\") - 1) & "\TerminalSubNumber\"
    Set rs = New ADODB.Recordset

    Dim mysql(0) As String
    
    mysql(0) = " SELECT * " & _
          " FROM " & Replace(���i�i��, " ", "") & ".txt"
    
    For a = 0 To UBound(mysql)
        'SQL���J��
        'cn.Execute mySQL(0)
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        If a = 0 Then ReDim RAN(rs.Fields.count, 0): j = 0
        
        'Sheets("Sheet1").Cells.ClearContents
        Do Until rs.EOF
            ReDim Preserve RAN(rs.Fields.count, j)
            For i = 0 To rs.Fields.count - 1
                'Sheets("Sheet1").Cells(J + 1, i + 1) = rs(i).Value
                RAN(i, j) = rs(i).Value
            Next i
            j = j + 1
            'Range("a2").CopyFromRecordset rs
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Function
Sub SQL_�}���}�ύX(���i�i��RAN, �}���}�ύXRAN, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).ActiveSheet
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    Dim mysql(0) As String
    
    mysql(0) = "SELECT "
    For i = 1 To ���i�i��RANc
        mysql(0) = mysql(0) & "[" & ���i�i��RAN(1, i - 1) & "],"
    Next i
    
    ReDim �d���ꗗRAN(���i�i��RANc + 9, 0)
    ReDim �[���ꗗran(0)
    mysql(0) = mysql(0) & "�\��_,�n�_����H����, �I�_����H����, �n�_���[�����ʎq, �I�_���[�����ʎq,�n�_���L���r�e�B,�I�_���L���r�e�B,����_,������_ ,RLTFtoPVSW_,���l_" & _
          " FROM myTable " & _
          " WHERE " & "[RLTFtoPVSW_]='Found'" '& _
          " AND [������_] IS NOT NULL"

    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim j As Long: j = 0
        Dim jj As Long: jj = 0
        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            '���type�̑Ώۂɂ��邩�m�F
            findFlg = False
            For i = 1 To ���i�i��RANc
                If Not IsNull(rs(i - 1)) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            
            If findFlg = False Then
                GoTo line20
            End If
            
            '�ǉ�
            ReDim Preserve �d���ꗗRAN(���i�i��RANc + 9, j)
            
            For i = 1 To ���i�i��RANc
                �d���ꗗRAN(i - 1, j + 0) = rs(i - 1)
            Next i
                '�n�_
                �d���ꗗRAN(���i�i��RANc + 0, j + 0) = rs(���i�i��RANc + 0) '�\��
                �d���ꗗRAN(���i�i��RANc + 1, j + 0) = rs(���i�i��RANc + 1) '��
                �d���ꗗRAN(���i�i��RANc + 2, j + 0) = rs(���i�i��RANc + 2)
                �d���ꗗRAN(���i�i��RANc + 3, j + 0) = rs(���i�i��RANc + 3) '�[��
                �d���ꗗRAN(���i�i��RANc + 4, j + 0) = rs(���i�i��RANc + 4)
                �d���ꗗRAN(���i�i��RANc + 5, j + 0) = rs(���i�i��RANc + 5) 'cav
                �d���ꗗRAN(���i�i��RANc + 6, j + 0) = rs(���i�i��RANc + 6)
                �d���ꗗRAN(���i�i��RANc + 7, j + 0) = rs(���i�i��RANc + 7) '����_
                �d���ꗗRAN(���i�i��RANc + 8, j + 0) = rs(���i�i��RANc + 8) '������_
                �d���ꗗRAN(���i�i��RANc + 9, j + 0) = rs(���i�i��RANc + 10) '���l_
                
            '�n�_�[���������ǉ�
            For i = LBound(�[���ꗗran) To UBound(�[���ꗗran)
                findFlg = False
                If �[���ꗗran(i) = rs(���i�i��RANc + 3) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            If findFlg = False Then
                ReDim Preserve �[���ꗗran(jj)
                �[���ꗗran(jj) = rs(���i�i��RANc + 3)
                jj = jj + 1
            End If
            '�I�_�[���������ǉ�
            For i = LBound(�[���ꗗran) To UBound(�[���ꗗran)
                findFlg = False
                If �[���ꗗran(i) = rs(���i�i��RANc + 4) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            If findFlg = False Then
                ReDim Preserve �[���ꗗran(jj)
                �[���ꗗran(jj) = rs(���i�i��RANc + 4)
                jj = jj + 1
            End If
            j = j + 1
line20:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_�݊��[��cav(�݊��[��cavRAN, �݊��[��RAN, ���i�i��str, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
    cn.Open
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim �݊��[��cavRAN(2, 0)
    Dim mysql(1) As String, ����(4) As String
    '�n�_���̉�H
    
    mysql(0) = " SELECT �n�_���[�����ʎq,�n�_���L���r�e�B" & _
          " FROM �͈� " & _
          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null and �n�_���[�����ʎq <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" '& _
          " GROUP BY �n�_���[�����ʎq,�n�_���L���r�e�B"
    '�I�_���̉�H
    mysql(1) = " SELECT �I�_���[�����ʎq,�I�_���L���r�e�B" & _
          " FROM �͈� " & _
          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null and �I�_���[�����ʎq <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" '& _
          " GROUP BY �I�_���[�����ʎq,�I�_���L���r�e�B"
    Dim cnt As Long
    j = 0
    For a = 0 To 1
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        Do Until rs.EOF
            ReDim Preserve �݊��[��cavRAN(2, j)
            
            For p = 0 To rs.Fields.count - 1
                �݊��[��cavRAN(p, j) = rs(p)
            Next p
            For i = LBound(�݊��[��RAN, 2) To UBound(�݊��[��RAN, 2)
                If �݊��[��RAN(0, i) = rs(0) Then
                    �݊��[��cavRAN(2, j) = �݊��[��RAN(1, i)
                    Exit For
                End If
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Sub
Sub SQL_�݊��[��cav_1998(�݊��[��cavRAN, �݊��[��RAN, ���i�i��str, myBookName)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
    cn.Open
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim �݊��[��cavRAN(5, 0)
    Dim mysql(0) As String, ����(4) As String
    '�n�_���̉�H
    
    mysql(0) = " SELECT �n�_���[�����ʎq,�n�_���L���r�e�B,�I�_���[�����ʎq,�I�_���L���r�e�B" & _
          " FROM �͈� " & _
          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" '& _
          " GROUP BY �n�_���[�����ʎq,�n�_���L���r�e�B"
    '�I�_���̉�H
'    mySQL(1) = " SELECT �I�_���[�����ʎq,�I�_���L���r�e�B" & _
'          " FROM �͈� " & _
'          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " <> Null and �I�_���[�����ʎq <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" '& _
          " GROUP BY �I�_���[�����ʎq,�I�_���L���r�e�B"
          
    Dim cnt As Long
    j = 0
    For a = 0 To 0
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        Do Until rs.EOF
            ReDim Preserve �݊��[��cavRAN(5, j)
            
            For p = 0 To rs.Fields.count - 1
                �݊��[��cavRAN(p, j) = rs(p)
            Next p
            
            Dim �n�_flg As Boolean: �n�_flg = False
            Dim �I�_flg As Boolean: �I�_flg = False
            For i = LBound(�݊��[��RAN, 2) To UBound(�݊��[��RAN, 2)
                '�n�__�[���������Ȃ�����W���Z�b�g
                If �݊��[��RAN(0, i) = rs(0) Then
                    �݊��[��cavRAN(4, j) = �݊��[��RAN(1, i)
                    �n�_flg = True
                End If
                '�I�__
                If �݊��[��RAN(0, i) = rs(2) Then
                    �݊��[��cavRAN(5, j) = �݊��[��RAN(1, i)
                    �I�_flg = True
                End If
                If �n�_flg = True And �I�_flg = True Then Exit For
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Sub

Sub SQL_�z����H�擾(�z����HRAN, ���i�i��str, �T�ustr)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
    cn.Open
    
    Dim �R�����g As String: �R�����g = "RLTFtoPVSW_" & " = " & "Found"
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    Set rs = New ADODB.Recordset
    
    Dim mysql As String
    mysql = " SELECT �F��_,�T��_,�n�_���[�����ʎq,�n�_���}_,�n�_���n��,�I�_���[�����ʎq,�I�_���}_,�I�_���n��" & _
          " FROM �͈� " & _
          " WHERE " & Chr(34) & ���i�i��str & Chr(34) & " = " & �T�ustr & " AND " & "RLTFtoPVSW_='Found'"   '& _
          " GROUP BY  �n�_���[�����ʎq,�I�_���[�����ʎq"

    'SQL���J��
    rs.Open mysql, cn, adOpenStatic
    '�z��Ɋi�[
    ReDim �z����HRAN(rs.Fields.count - 1, rs.RecordCount - 1)
    Do Until rs.EOF
        For p = 0 To rs.Fields.count - 1
            �z����HRAN(p, j) = rs(p)
        Next p
        j = j + 1
        rs.MoveNext
    Loop
    rs.Close
    cn.Close

End Sub
Sub SQL_�z���T�u�擾(�z���T�uRAN, ���i�i��str)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("�[���ꗗ")
        Dim myRange As Range: Set myRange = .Cells.Find("�[�����i��", , , 1)
        Dim firstRow As Long: firstRow = myRange.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, myRange.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(myRange.Row, .Columns.count).End(xlToLeft).Column
        Set myRange = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    Set rs = New ADODB.Recordset
    
    Dim mysql As String
    mysql = " SELECT [" & ���i�i��str & "] " & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
          " GROUP BY [" & ���i�i��str & "]" & _
          " ORDER BY len([" & ���i�i��str & "]),[" & ���i�i��str & "]" ' & _
          " AND " & "RLTFtoPVSW_='Found'"   '& _
          " GROUP BY  �n�_���[�����ʎq,�I�_���[�����ʎq"

    'SQL���J��
    rs.Open mysql, cn, adOpenStatic
    '�z��Ɋi�[
    ReDim �z���T�uRAN(rs.Fields.count - 1, rs.RecordCount - 1)
    Do Until rs.EOF
        For p = 0 To rs.Fields.count - 1
            �z���T�uRAN(p, j) = rs(p)
        Next p
        j = j + 1
        rs.MoveNext
    Loop
    
    ReDim Preserve �z���T�uRAN(0, UBound(�z���T�uRAN, 2) + 1)
    �z���T�uRAN(0, UBound(�z���T�uRAN, 2)) = "Base"
    rs.Close
    cn.Close

End Sub

Sub SQL_�z��_�[���o�H�擾(�[���o�HRAN, ���i�i��str, �[��str)
      
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("PVSW_RLTF")
        Dim �d�����ʖ� As Range: Set �d�����ʖ� = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = �d�����ʖ�.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(�d�����ʖ�.Row, .Columns.count).End(xlToLeft).Column
        Set �d�����ʖ� = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    ReDim �[���o�HRAN(6, 0)
    Dim mysql(1) As String, ����(4) As String
    '�n�_���̉�H
    mysql(0) = " SELECT �n�_���[�����ʎq,�I�_���[�����ʎq, �T��_,�F��_,�I�_���}_,�d�㐡�@_,����_" & _
          " FROM �͈� " & _
          " WHERE [�n�_���[�����ʎq] = '" & �[��str & "'" & _
          " AND " & "RLTFtoPVSW_='Found'" & " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """""
    '�I�_���̉�H
    mysql(1) = " SELECT �I�_���[�����ʎq,�n�_���[�����ʎq, �T��_,�F��_,�n�_���}_,�d�㐡�@_,����_" & _
          " FROM �͈� " & _
          " WHERE [�I�_���[�����ʎq] = '" & �[��str & "'" & _
          " AND " & "RLTFtoPVSW_='Found'" & " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """""
    For a = 0 To 1
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        
        Do Until rs.EOF
            ReDim Preserve �[���o�HRAN(rs.Fields.count - 1, j)
            For p = 0 To rs.Fields.count - 1
                �[���o�HRAN(p, j) = rs(p)
            Next p
            j = j + 1
            rs.MoveNext
        Loop
        
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_���i�ʒ[���ꗗ_�g�p�d���m�F(�g�p�d��ran, ���i�i��str)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("PVSW_RLTF")
        Dim myRange As Range: Set myRange = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = myRange.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, myRange.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(myRange.Row, .Columns.count).End(xlToLeft).Column
        Set myRange = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    Set rs = New ADODB.Recordset
    Dim mysql(1) As String
    'Dim �g�p�d��ran()
    ReDim �g�p�d��ran(3, 0)
    j = 0
    mysql(0) = " SELECT [" & ���i�i��str & "],[�n�_���[�����ʎq] ,[�n�_���[�����i��],[�n�_���L���r�e�B]" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
          " AND " & "RLTFtoPVSW_='Found'"   '& _
          " GROUP BY  �n�_���[�����ʎq,�I�_���[�����ʎq"
    mysql(1) = " SELECT [" & ���i�i��str & "] ,[�I�_���[�����ʎq],[�I�_���[�����i��],[�I�_���L���r�e�B]" & _
          " FROM �͈� " & _
          " WHERE [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
          " AND " & "RLTFtoPVSW_='Found'"
    For a = LBound(mysql) To UBound(mysql)
          '& _
              " GROUP BY  �I�_���[�����ʎq,�I�_���[�����ʎq"
        'SQL���J��
        rs.Open mysql(a), cn, adOpenStatic
        '�g�p���Ă���CAV���i�[
        Do Until rs.EOF
            ReDim Preserve �g�p�d��ran(rs.Fields.count - 1, j)
            For p = 0 To rs.Fields.count - 1
                �g�p�d��ran(p, j) = rs(p)
            Next p
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_�z���}�p_���i�i��_�\��_SUB(RAN, ���i�i��str, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(1, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
    For k = 0 To 0
        mysql(0) = " SELECT [" & ���i�i��str & "],�\��_,'" & ���i�i��str & "'" & _
              " FROM �͈� " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
              " ORDER BY [" & ���i�i��str & "] ASC"
'        mySQL(1) = " SELECT [" & ���i�i��str & "],�I�_���[�����i��,�I�_���[�����ʎq,TI2,'" & ���i�i��str & "'" & _
'              " FROM �͈� " & _
'              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
'              " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
'              " ORDER BY [" & ���i�i��str & "] ASC"
    
        'SQL���J��
        rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            flg = False
            '�ǉ�
            If flg = False Then
                If rs(1) & rs(0) <> "" Then
                    j = j + 1
                    ReDim Preserve RAN(1, j)
                    RAN(0, j) = Replace(rs(2), " ", "") & "_" & rs(1) '���i�i��_�\��
                    RAN(1, j) = rs(0) 'Sub
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next k
    cn.Close

End Sub

Public Function SQL_�d�����RANset(RAN, ���i�i��str, myBook, �[��)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF���[")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With

    SQL_�d�����RANset = 0
    ReDim RAN(9, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
    For k = 0 To 0
        mysql(0) = " SELECT [" & ���i�i��str & "],�[�����i��,�T��_,�ؒf��_,�[�q_,�[�����ʎq,�\��_,�}_,����_,'" & ���i�i��str & "'" & _
              " FROM �͈� " & _
              " WHERE " & "[RLTFtoPVSW_]='Found' AND [�[�����ʎq]='" & �[�� & "'" & _
              " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
              " ORDER BY �ؒf��_ DESC, ����_ ASC"  '�����̃\�[�g�����ĂȂ�
'        mySQL(1) = " SELECT [" & ���i�i��str & "],�I�_���[�����i��,�I�_���[�����ʎq,TI2,'" & ���i�i��str & "'" & _
'              " FROM �͈� " & _
'              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
'              " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
'              " ORDER BY [" & ���i�i��str & "] ASC"
        'SQL���J��
        rs.CursorLocation = adUseClient
        rs.Open mysql(k), cn, adOpenKeyset, adLockOptimistic, 512
        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            flg = False
            If flg = False Then
                If rs(1) & rs(0) <> "" Then
                    If Left(rs(4), 4) = "7409" Then
                        ReDim Preserve RAN(9, j)
                        For i = 0 To rs.Fields.count - 1
                            RAN(i, j) = Replace(rs(i), " ", "")
                        Next i
                        j = j + 1
                        SQL_�d�����RANset = j
                    End If
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next k
    cn.Close
    
    If j > 0 Then
        '�ؒf��_���ŕ��ёւ�����
        Dim myAry1()
        myAry1 = WorksheetFunction.Transpose(RAN) 'SQL�ŃZ�b�g�����z������ւ���
        '2�����o�u���\�[�g
        Call BubbleSort2(myAry1, 4) '����
        RAN = WorksheetFunction.Transpose(myAry1)
    End If
End Function

Sub SQL_�z���}�p_��H(RAN, ���i�i��str, myBook)
    
    '�Q�Ɛݒ�= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("�d�����ʖ�", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "�͈�"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim RAN(11, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
    For k = 0 To 0
        mysql(0) = " SELECT [" & ���i�i��str & "],�\��_,�F��_,�n�_���[�����ʎq,�I�_���[�����ʎq,'" & ���i�i��str & "'" & ",�n�_���n��,�n�_���L���r�e�B,�I�_���n��,�I�_���L���r�e�B" & _
              " FROM �͈� " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
              " ORDER BY [" & ���i�i��str & "] ASC"
'        mySQL(1) = " SELECT [" & ���i�i��str & "],�I�_���[�����i��,�I�_���[�����ʎq,TI2,'" & ���i�i��str & "'" & _
'              " FROM �͈� " & _
'              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
'              " AND [" & ���i�i��str & "] IS NOT NULL AND [" & ���i�i��str & "] <> """"" & _
'              " ORDER BY [" & ���i�i��str & "] ASC"
    
        'SQL���J��
        rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
        If rs(0).Type <> 202 Then Stop 'rs�̐��i�i��str�̃f�[�^�^�C�v��202����Ȃ����當���񂪔�����
                                       'PVSW_RLTF�̏����ݒ��@�ɂ���Ƃ�
        Do Until rs.EOF
            flg = False
            '�ǉ�
            If flg = False Then
                If rs(1) & rs(0) <> "" Then
                    j = j + 1
                    ReDim Preserve RAN(11, j)
                    RAN(0, j) = Replace(rs(5), " ", "") '���i�i��
                    RAN(1, j) = rs(0) 'Sub
                    RAN(2, j) = rs(1)
                    RAN(3, j) = rs(2)
                    RAN(4, j) = rs(3)
                    RAN(5, j) = rs(4)
                    RAN(6, j) = rs(5)
                    RAN(7, j) = rs(6)
                    RAN(8, j) = rs(7)
                    RAN(9, j) = rs(8)
                    RAN(10, j) = rs(9)
                    RAN(11, j) = �F�ϊ�(rs(2), clocode1, clocode2, clofont) '�F��long
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next k
    cn.Close

End Sub

Sub SQL���ǂ�_��n����Ǝ�(RAN, ���i�i��str)
    
    '�V�[�g�����傫���V�[�g�̌���
    Dim wsTemp As Worksheet, wsNumber As Long
    For Each wsTemp In wb(3).Worksheets
        If IsNumeric(wsTemp.Name) Then
            If CLng(wsTemp.Name) > wsNumber Then
                wsNumber = wsTemp.Name
            End If
        End If
    Next wsTemp
    
    If wsNumber = 0 Then
        MsgBox "�V�[�g���ɐ�����������܂���B���f���܂�"
        Call �œK�����ǂ�
        wb(3).Close
        End
    End If
    
    With wb(3).Sheets(CStr(wsNumber))
        Dim myKey As Range: Set myKey = .Cells.Find("key_", , , 1)
        Dim firstRow As Long: firstRow = myKey.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        Dim koseiRow As Long
        koseiRow = .Columns(myKey.Column).Find("CONP No", , , 1).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(firstRow, 2), .Cells(lastRow, lastCol)).Name = "�͈�"
    End With
    
    With wb(3).Sheets(CStr(wsNumber))
        Dim Col0 As Long: Col0 = .Rows(firstRow).Find("key_", , , 1).Column
        Dim Col1 As Long: Col1 = .Rows(firstRow).Find(���i�i��str, , , 1).Column
        ReDim RAN(1, 0)
        c = 0
        For Y = koseiRow + 1 To lastRow
            �\�� = .Cells(Y, Col0)
            If �\�� <> "" Then
                ��n����Ǝ�str = .Cells(Y, Col1)
                ReDim Preserve RAN(1, UBound(RAN, 2) + 1)
                RAN(0, UBound(RAN, 2)) = �\��
                RAN(1, UBound(RAN, 2)) = ��n����Ǝ�str
            End If
        Next Y
        Set myKey = Nothing
    End With

    ��n����Ǝ҃V�[�g�� = wsNumber & "��"
End Sub


