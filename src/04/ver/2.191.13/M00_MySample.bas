Attribute VB_Name = "M00_MySample"
Sub ADO��SQL�J��(RAN, myBook As Workbook, ���i�i��str As String)
    
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
        '[���i�i��]���猩��[PVSW_RLTF]�Ƀ��C���i�Ԃ��������A�������΂�
        For k = 0 To 1
            mysql(0) = " SELECT �͈�b.[�ȈՃ|�C���g],�͈�a.[�n�_����H����],�͈�a.[�F_],�͈�a.[�F��_]" & _
                  " FROM �͈�a INNER JOIN �͈�b" & _
                  " ON �͈�a.[�n�_���[�����ʎq] = �͈�b.[�[����] And �͈�a.[�n�_���[�����i��] = �͈�b.[�[�����i��] AND �͈�a.[�n�_���L���r�e�B] = �͈�b.[Cav] " & _
                  " WHERE " & "�͈�a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND �͈�a.[" & ���i�i��str & "] IS NOT NULL AND �͈�a.[" & ���i�i��str & "] <> """""
        
            mysql(0) = " SELECT �͈�a.* ,�͈�b.*" & _
                  " FROM �͈�a INNER JOIN �͈�b" & _
                  " ON �͈�a.�n�_���[�����ʎq = �͈�b.�[���� And �͈�a.�n�_���[�����i�� = �͈�b.�[�����i�� AND �͈�a.�n�_���L���r�e�B = �͈�b.Cav " & _
                  " WHERE " & "�͈�a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND �͈�a.[" & ���i�i��str & "] IS NOT NULL AND �͈�a.[" & ���i�i��str & "] <> """""
                  
            mysql(1) = " SELECT �͈�b.�ȈՃ|�C���g,�͈�a.�I�_����H����,�͈�a.�F_,�͈�a.�F��_" & _
                  " FROM �͈�a INNER JOIN �͈�b" & _
                  " ON �͈�a.�I�_���[�����ʎq = �͈�b.�[���� And �͈�a.�I�_���[�����i�� = �͈�b.�[�����i�� AND �͈�a.�I�_���L���r�e�B = �͈�b.Cav " & _
                  " WHERE " & "�͈�a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND �͈�a.[" & ���i�i��str & "] IS NOT NULL AND �͈�a.[" & ���i�i��str & "] <> """""
                  
            'SQL���J��
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
          Stop
            
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
                                    flg = True
                                End If
                            End If
                        End If
                    End If
                Next r
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

