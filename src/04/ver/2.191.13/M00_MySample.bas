Attribute VB_Name = "M00_MySample"
Sub ADOでSQL開く(RAN, myBook As Workbook, 製品品番str As String)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Properties("Jet OLEDB:Engine Type") = 35 'これで指定できてない,37だと型が一致しないエラー
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Debug.Print "Jet OLEDB:Engine Type", cn.Properties("Jet OLEDB:Engine Type")
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲a"
    End With
    
    With myBook.Sheets("ポイント一覧")
        Set key = .Cells.Find("端末矢崎品番", , , 1)
        firstRow = key.Row
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(firstRow, key.Column), .Cells(lastRow, lastCol)).Name = "範囲b"
        Set key = Nothing
    End With
    
    ReDim RAN(3, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
        '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
        For k = 0 To 1
            mysql(0) = " SELECT 範囲b.[簡易ポイント],範囲a.[始点側回路符号],範囲a.[色_],範囲a.[色呼_]" & _
                  " FROM 範囲a INNER JOIN 範囲b" & _
                  " ON 範囲a.[始点側端末識別子] = 範囲b.[端末№] And 範囲a.[始点側端末矢崎品番] = 範囲b.[端末矢崎品番] AND 範囲a.[始点側キャビティ] = 範囲b.[Cav] " & _
                  " WHERE " & "範囲a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND 範囲a.[" & 製品品番str & "] IS NOT NULL AND 範囲a.[" & 製品品番str & "] <> """""
        
            mysql(0) = " SELECT 範囲a.* ,範囲b.*" & _
                  " FROM 範囲a INNER JOIN 範囲b" & _
                  " ON 範囲a.始点側端末識別子 = 範囲b.端末№ And 範囲a.始点側端末矢崎品番 = 範囲b.端末矢崎品番 AND 範囲a.始点側キャビティ = 範囲b.Cav " & _
                  " WHERE " & "範囲a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND 範囲a.[" & 製品品番str & "] IS NOT NULL AND 範囲a.[" & 製品品番str & "] <> """""
                  
            mysql(1) = " SELECT 範囲b.簡易ポイント,範囲a.終点側回路符号,範囲a.色_,範囲a.色呼_" & _
                  " FROM 範囲a INNER JOIN 範囲b" & _
                  " ON 範囲a.終点側端末識別子 = 範囲b.端末№ And 範囲a.終点側端末矢崎品番 = 範囲b.端末矢崎品番 AND 範囲a.終点側キャビティ = 範囲b.Cav " & _
                  " WHERE " & "範囲a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND 範囲a.[" & 製品品番str & "] IS NOT NULL AND 範囲a.[" & 製品品番str & "] <> """""
                  
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
          Stop
            
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                '登録があるか確認
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
                '追加
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

