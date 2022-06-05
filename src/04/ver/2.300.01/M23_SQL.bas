Attribute VB_Name = "M23_SQL"

Sub SQL_配索端末取得(配索端末RAN, 製品品番str, サブstr)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim 配索端末RAN(1, 0)
    
    Dim mysql(1) As String, 条件(1) As String
    '始点側の回路
    mysql(0) = " SELECT [始点側端末識別子], [始点側ハメ],[色呼_],[構成_],[サ呼_],[RLTFtoPVSW_],[始点側マ_]" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] = '" & サブstr & "'" & _
          " AND [始点側端末識別子] IS NOT NULL"  ' & _
          " GROUP BY  [始点側端末識別子],[始点側ハメ],[色呼_],[始点側マ_]"
          
    '終点側の回路
    mysql(1) = " SELECT 終点側端末識別子, 終点側ハメ,色呼_,構成_,サ呼_,RLTFtoPVSW_,終点側マ_" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] = '" & サブstr & "'" & _
          " AND 終点側端末識別子 IS NOT NULL " '& _
          " GROUP BY  終点側端末識別子,終点側ハメ,色呼_,終点側マ_"
    
    For a = 0 To 1
        'SQLを開く=ここでエラーになる時、もしかしてPVSW_RLTFで全部のセルエンター実行せないかんかも
        rs.Open mysql(a), cn, adOpenStatic
        Debug.Print rs.RecordCount
        '配列に格納
        Do Until rs.EOF
            If rs(1).Value = "後" Then
                条件(0) = rs(0).Value
                条件(1) = rs(2).Value
                '条件(2) = rs(3).Value
            Else
                条件(0) = rs(0).Value
                条件(1) = ""
                '条件(2) = rs(3).Value
            End If
            If rs(0).Value = "" Then GoTo line10
            For p = 0 To UBound(配索端末RAN, 2)
                If 配索端末RAN(0, p) = 条件(0) Then
                    If 配索端末RAN(1, p) = 条件(1) Then
                        GoTo line10
                    End If
                End If
            Next p
            '無いので格納
            ReDim Preserve 配索端末RAN(1, UBound(配索端末RAN, 2) + j)
            For i = 0 To 1
                配索端末RAN(i, UBound(配索端末RAN, 2)) = 条件(i)
            Next i
            j = 1
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Sub

Public Function SQL_自動機(自動機RAN)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    Call checkSheet("PVSW_RLTF", wb(0), True, True)
    
    With Sheets("PVSW_RLTF")
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim 自動機RAN(0, 0)
    
    Dim mysql(0) As String, 条件(1) As String
    '始点側の回路
    mysql(0) = " SELECT [自動機]" & _
          " FROM 範囲 " & _
          " WHERE [自動機] IS NOT NULL" & _
          " GROUP BY [自動機]"
          
    For a = 0 To 0
        'SQLを開く=ここでエラーになる時、もしかしてPVSW_RLTFで全部のセルエンター実行せないかんかも
        rs.Open mysql(a), cn, adOpenStatic
        Debug.Print rs.RecordCount
        j = 0
        '配列に格納
        Do Until rs.EOF
            '無いので格納
            ReDim Preserve 自動機RAN(0, j)
            自動機RAN(0, j) = rs(0)
            j = j + 1
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Function


Sub SQL_配索端末取得_端末用端末(配索端末RAN, 端末str)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim 配索端末RAN(0)
    
    Dim mysql(1) As String, 条件(1) As String
    '始点側の回路
    mysql(0) = " SELECT [終点側端末識別子]" & _
          " FROM 範囲 " & _
          " WHERE [始点側端末識別子] = '" & 端末str & "'" & _
          " AND [終点側端末識別子] IS NOT NULL" & _
          " GROUP BY [終点側端末識別子]"
          
    '終点側の回路
    mysql(1) = " SELECT [始点側端末識別子]" & _
          " FROM 範囲 " & _
          " WHERE [終点側端末識別子] = '" & 端末str & "'" & _
          " AND [始点側端末識別子] IS NOT NULL" & _
          " GROUP BY [始点側端末識別子]"
    
    For a = 0 To 1
        'SQLを開く=ここでエラーになる時、もしかしてPVSW_RLTFで全部のセルエンター実行せないかんかも
        rs.Open mysql(a), cn, adOpenStatic
        Debug.Print rs.RecordCount
        '配列に格納
        Do Until rs.EOF
            For i = LBound(配索端末RAN) To UBound(配索端末RAN)
                If rs(0) = 配索端末RAN(i) Then GoTo line10
            Next i
            '無いので格納
            ReDim Preserve 配索端末RAN(UBound(配索端末RAN) + j)
            配索端末RAN(UBound(配索端末RAN)) = rs(0)
            j = 1
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Sub


Sub SQL_配索端末取得_端末用回路(ran, 端末v, 端末str)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim ran(6, 0)
    
    Dim mysql(1) As String, 条件(1) As String
    '始点側の回路
    mysql(0) = " SELECT [始点側端末識別子],[終点側端末識別子],[色呼_],[構成_],[サ呼_],[RLTFtoPVSW_],[終点側マ_]" & _
          " FROM 範囲 " & _
          " WHERE [始点側端末識別子] =" & "'" & 端末v & "'" & _
          " AND [終点側端末識別子] =" & "'" & 端末str & "'" '& _
          " AND [始点側端末識別子] IS NOT NULL"  ' & _
          " GROUP BY  [始点側端末識別子],[始点側ハメ],[色呼_],[始点側マ_]"
          
    '終点側の回路
    mysql(1) = " SELECT [終点側端末識別子], [始点側端末識別子],[色呼_],[構成_],[サ呼_],[RLTFtoPVSW_],[始点側マ_]" & _
          " FROM 範囲 " & _
          " WHERE [終点側端末識別子] =" & "'" & 端末v & "'" & _
          " AND [始点側端末識別子] =" & "'" & 端末str & "'" '& _
          " AND [終点側端末識別子] IS NOT NULL " '& _
          " GROUP BY  終点側端末識別子,終点側ハメ,色呼_,終点側マ_"
    
    For a = 0 To 1
        'SQLを開く=ここでエラーになる時、もしかしてPVSW_RLTFで全部のセルエンター実行せないかんかも
        rs.Open mysql(a), cn, adOpenStatic
        '配列に格納
        Do Until rs.EOF
            '同じ構成№は格納しない
            For i = 0 To UBound(ran, 2)
                If ran(3, i) = rs(3) Then GoTo line20
            Next i
            '格納
            ReDim Preserve ran(6, UBound(ran, 2) + j)
            For i = 0 To UBound(ran, 1)
                ran(i, UBound(ran, 2)) = rs(i)
            Next i
            j = 1
line20:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Sub


Sub SQL_配索端末取得2(配索端末RAN, 製品品番str, サブstr)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    Set rs = New ADODB.Recordset
    ReDim 配索端末RAN(1, 0)
    
    Dim mysql(1) As String, 条件(1) As String
    '始点側の回路
    mysql(0) = " SELECT [始点側端末識別子], [始点側ハメ],[色呼_],[構成_],[サ呼_],[RLTFtoPVSW_],[始点側マ_]" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] = '" & サブstr & "'" & _
          " AND [始点側端末識別子] IS NOT NULL"  ' & _
          " GROUP BY  [始点側端末識別子],[始点側ハメ],[色呼_],[始点側マ_]"
          
    '終点側の回路
    mysql(1) = " SELECT 終点側端末識別子, 終点側ハメ,色呼_,構成_,サ呼_,RLTFtoPVSW_,終点側マ_" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] = '" & サブstr & "'" & _
          " AND 終点側端末識別子 IS NOT NULL " '& _
          " GROUP BY  終点側端末識別子,終点側ハメ,色呼_,終点側マ_"
    
    For a = 0 To 1
        'SQLを開く=ここでエラーになる時、もしかしてPVSW_RLTFで全部のセルエンター実行せないかんかも
        rs.Open mysql(a), cn, adOpenStatic
        Debug.Print rs.RecordCount
        '配列に格納
        Do Until rs.EOF
            If rs(1).Value = "後" Then
                条件(0) = rs(0).Value
                条件(1) = rs(2).Value
                '条件(2) = rs(3).Value
            Else
                条件(0) = rs(0).Value
                条件(1) = ""
                '条件(2) = rs(3).Value
            End If
            If rs(0).Value = "" Then GoTo line10
            For p = 0 To UBound(配索端末RAN, 2)
                If 配索端末RAN(0, p) = 条件(0) Then
                    GoTo line10
                End If
            Next p
            '無いので格納
            ReDim Preserve 配索端末RAN(1, UBound(配索端末RAN, 2) + j)
            For i = 0 To 1
                配索端末RAN(i, UBound(配索端末RAN, 2)) = 条件(i)
            Next i
            j = 1
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next a
    
    cn.Close

End Sub


Public Function SQL_配索図_端末一覧(myBookName, 冶具type)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("端末一覧")
        Dim myKey As Range: Set myKey = .Cells.Find("端末矢崎品番", , , 1)
        Dim firstRow As Long: firstRow = myKey.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        Set myKey = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myArea"
    End With
    
    Set rs = New ADODB.Recordset
    Dim mysql(0) As String
    
    'この冶具で使用する端末一覧を配列にセット
    ReDim 端末一覧ran(0)
    For r = LBound(製品品番Ran, 2) To UBound(製品品番Ran, 2)
        If 製品品番Ran(製品品番RAN_read(製品品番Ran, "結き"), r) = 冶具type Then
            製品品番str = 製品品番Ran(製品品番RAN_read(製品品番Ran, "メイン品番"), r)
            mysql(0) = " SELECT [端末№]" & _
          " FROM myArea " & _
          " WHERE [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """""
            For a = LBound(mysql) To UBound(mysql)
                'SQLを開く=ここでエラーになる時、もしかしてPVSW_RLTFで全部のセルエンター実行せないかんかも
                rs.Open mysql(a), cn, adOpenStatic
                'Debug.Print rs.RecordCount
                '配列に格納
                Do Until rs.EOF
                    For p = 0 To UBound(端末一覧ran, 1)
                        If 端末一覧ran(p) = rs(0) Then
                            GoTo line10 'あるので次のレコード
                        End If
                    Next p
                    '無いので格納
                    ReDim Preserve 端末一覧ran(UBound(端末一覧ran, 1) + j)
                    端末一覧ran(UBound(端末一覧ran)) = rs(0)
                    j = 1
line10:
                    rs.MoveNext
                Loop
                rs.Close
            Next a
        End If
    Next r
    cn.Close
    
    'このシートに端末があるか確認
    ReDim 端末無い一覧RAN(0): j = 0
    For p = 0 To UBound(端末一覧ran)
        
        Set myfnd = ActiveSheet.Cells.Find(端末一覧ran(p), , , 1)
        If myfnd Is Nothing Then
            ReDim Preserve 端末無い一覧RAN(UBound(端末無い一覧RAN) + j)
            端末無い一覧RAN(UBound(端末無い一覧RAN)) = 端末一覧ran(p)
            j = 1
        End If
    Next p
    SQL_配索図_端末一覧 = 端末無い一覧RAN
End Function


Sub SQL_配索後ハメ取得(配索後ハメRAN, 製品品番str, サブstr)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim 配索後ハメRAN(5, 0)
    Dim mysql(1) As String, 条件(4) As String
    '始点側の回路
    mysql(0) = " SELECT 色呼_, サ呼_,始点側端末識別子,始点側マ_,始点側ハメ,生区_" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] = '" & サブstr & "'" & _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " AND " & "始点側ハメ = '後'"
    '終点側の回路
    mysql(1) = " SELECT 色呼_, サ呼_,終点側端末識別子,終点側マ_,終点側ハメ,生区_" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] = '" & サブstr & "'" & _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " AND " & "終点側ハメ = '後'"
    For a = 0 To 1
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        
        Do Until rs.EOF
            ReDim Preserve 配索後ハメRAN(rs.fields.count - 1, j)
            For p = 0 To rs.fields.count - 1
                配索後ハメRAN(p, j) = rs(p)
            Next p
            j = j + 1
            rs.MoveNext
        Loop
        
        rs.Close
    Next a
    cn.Close

End Sub
Sub SQL_配索後ハメ点滅取得(ran, 製品品番str)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim ran(2, 0)
    ReDim ran(4, 0)
    Dim mysql(1) As String, 条件(4) As String
    '始点側の回路
    mysql(0) = " SELECT 始点側端末識別子,始点側キャビティ,始点側ハメ,複ID_,生区_" & _
          " FROM 範囲 " & _
          " WHERE  RLTFtoPVSW_='Found'"

    '終点側の回路
    mysql(1) = " SELECT 終点側端末識別子,終点側キャビティ,終点側ハメ,複ID_,生区_" & _
          " FROM 範囲 " & _
          " WHERE  RLTFtoPVSW_='Found'"
    For a = 0 To 1
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        
        Do Until rs.EOF
            ReDim Preserve ran(rs.fields.count - 1, j)
            For p = 0 To rs.fields.count - 1
                ran(p, j) = rs(p)
            Next p
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub


Sub SQL_互換率算出(互換率RAN, 互換端末RAN, 製品品番str)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
    cn.Open
    Set rs = New ADODB.Recordset
    
    With Sheets("PVSW_RLTF")
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim 互換率RAN(5, 0)
    Dim mysql(0) As String, 条件(4) As String
    '始点側の回路
    mysql(0) = " SELECT 始点側端末識別子," & Chr(34) & "始点側キャビティ" & Chr(34) & ",終点側端末識別子," & Chr(34) & "終点側キャビティ" & Chr(34) & _
          " FROM 範囲 " & _
          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null " & _
          " AND " & "RLTFtoPVSW_='Found'"
    '終点側の回路
'    mySQL(1) = " SELECT 色呼_, サ呼_,終点側端末識別子,終点側マ_,終点側ハメ" & _
'          " FROM 範囲 " & _
'          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " = " & サブstr & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " AND " & "終点側ハメ = '後'"
    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        j = 0
        Do Until rs.EOF
            ReDim Preserve 互換率RAN(5, j)
            For p = 0 To rs.fields.count - 1
                互換率RAN(p, j) = rs(p)
            Next p
            For i = LBound(互換端末RAN, 2) To UBound(互換端末RAN, 2) '端末の座標を調べて登録
                If 互換率RAN(0, j) = 互換端末RAN(0, i) Then
                    互換率RAN(4, j) = 互換端末RAN(1, i)
                End If
                If 互換率RAN(2, j) = 互換端末RAN(0, i) Then
                    互換率RAN(5, j) = 互換端末RAN(1, i)
                End If
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_互換端末(互換端末RAN, 製品品番str, myBookName, 冶具type)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim 互換端末RAN(3, 0)
    Dim mysql(1) As String, 条件(4) As String
    '始点側の回路
    
    mysql(0) = " SELECT 始点側端末識別子 , COUNT(1)" & _
          " FROM 範囲 " & _
          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null and 始点側端末識別子 <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " GROUP BY 始点側端末識別子"
    '終点側の回路
    mysql(1) = " SELECT 終点側端末識別子 , COUNT(1)" & _
          " FROM 範囲 " & _
          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null and 終点側端末識別子 <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " GROUP BY 終点側端末識別子"
    j = 0
    For a = 0 To 1
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        Do Until rs.EOF
            For i = LBound(互換端末RAN, 2) To UBound(互換端末RAN, 2)
                If 互換端末RAN(0, i) = rs(0) Then
                    互換端末RAN(2, i) = 互換端末RAN(2, i) + rs(1) '端末№カウント
                    flg = 1
                End If
            Next i
            
            If flg = 0 Then '無い時は情報を追加
                ReDim Preserve 互換端末RAN(3, j)
                
                互換端末RAN(0, j) = rs(0) '端末№
                Set myfound = Workbooks(myBookName).Sheets("冶具_" & 冶具type).Cells.Find(rs(0), , , 1)
                If myfound Is Nothing Then '冶具座標
                    互換端末RAN(1, j) = "冶具座標無し"
                Else
                    互換端末RAN(1, j) = Workbooks(myBookName).Sheets("冶具_" & 冶具type).Cells.Find(rs(0), , , 1).Offset(, 1)
                End If
                互換端末RAN(2, j) = rs(1)
                j = j + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub
Public Function SQL_端末一覧(端末一覧ran, 製品品番str, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file 'ThisWorkbook.path & "\" & ThisWorkbook.Name
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With wb(0).Sheets("端末一覧")
        Dim 端末矢崎品番 As Range: Set 端末矢崎品番 = .Cells.Find("端末矢崎品番", , , 1)
        Dim firstRow As Long: firstRow = 端末矢崎品番.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 端末矢崎品番.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(端末矢崎品番.Row, .Columns.count).End(xlToLeft).Column
        lastCol = .Cells(firstRow, .Columns.count).End(xlToLeft).Column
        Set 端末矢崎品番 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim 端末一覧ran(3, 0)
    Dim mysql(0) As String
    '始点側の回路
    
    mysql(0) = " SELECT 端末矢崎品番 ,端末№, [" & 製品品番str & "],成型方向" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] is not Null AND [" & 製品品番str & "] <> """"" & _
          " ORDER BY [" & 製品品番str & "] ASC"  '& _
          " AND " & "RLTFtoPVSW_='Found'" & _
          " GROUP BY 始点側端末識別子"
    '終点側の回路
'    mySQL(1) = " SELECT 終点側端末識別子 , COUNT(1)" & _
'          " FROM 範囲 " & _
'          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null and 終点側端末識別子 <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " GROUP BY 終点側端末識別子"

    j = 0
    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        If rs(2).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
                                       
        Do Until rs.EOF
            ReDim Preserve 端末一覧ran(3, j)
            For i = LBound(端末一覧ran, 1) To UBound(端末一覧ran, 1)
                If IsNull(rs(i)) Then
                    端末一覧ran(i, j) = ""
                Else
                    端末一覧ran(i, j) = rs(i)
                End If
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
    SQL_端末一覧 = 端末一覧
End Function
Sub SQL_サブ端末数(サブ端末数RAN, 製品品番str, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file 'ThisWorkbook.path & "\" & ThisWorkbook.Name
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Workbooks(myBookName).Sheets("製品別端末一覧")
        Dim 端末矢崎品番 As Range: Set 端末矢崎品番 = .Cells.Find("端末矢崎品番", , , 1)
        Dim firstRow As Long: firstRow = 端末矢崎品番.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 端末矢崎品番.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(端末矢崎品番.Row, .Columns.count).End(xlToLeft).Column
        Set 端末矢崎品番 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim サブ端末数RAN(1, 0)
    Dim mysql(0) As String
    '始点側の回路
    
    mysql(0) = " SELECT [" & 製品品番str & "] ,COUNT(1)" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] is not Null AND [" & 製品品番str & "] <> """"" & _
          " GROUP BY [" & 製品品番str & "]" & _
          " ORDER BY [" & 製品品番str & "] ASC"
    '終点側の回路
'    mySQL(1) = " SELECT 終点側端末識別子 , COUNT(1)" & _
'          " FROM 範囲 " & _
'          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null and 終点側端末識別子 <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " GROUP BY 終点側端末識別子"

    j = 0
    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        If rs(0).Type <> 202 And rs(0).Type <> 200 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            ReDim Preserve サブ端末数RAN(1, j)
            
            For i = LBound(サブ端末数RAN, 1) To UBound(サブ端末数RAN, 1)
                サブ端末数RAN(i, j) = rs(i)
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Sub

Sub SQL_サブ端末数_動作確認用temp(サブ端末数RAN, 製品品番str, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file 'ThisWorkbook.path & "\" & ThisWorkbook.Name
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("製品別端末一覧")
        Dim key As Range: Set key = .Cells.Find("端末矢崎品番", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    'Dim 範囲 As DataTable
    
    ReDim サブ端末数RAN(1, 0)
    Dim mysql(0) As String
    '始点側の回路

    mysql(0) = " SELECT [" & 製品品番str & "]" & ",count(1)" & _
          " FROM myTable" & _
          " WHERE [" & 製品品番str & "] Is Not Null" & _
          " GROUP BY [" & 製品品番str & "]" & _
          " ORDER BY [" & 製品品番str & "] ASC"
          
    'mySQL(0) = " SELECT " & Chr(34) & 製品品番str & Chr(34) & " ,COUNT(1)" & _
          " FROM 範囲" & _
          " WHERE 端末矢崎品番 is not null" '& _
          " GROUP BY " & Chr(34) & 製品品番str & Chr(34) & _
          " ORDER BY " & Chr(34) & 製品品番str & Chr(34) & " ASC"
    '終点側の回路
'    mySQL(1) = " SELECT 終点側端末識別子 , COUNT(1)" & _
'          " FROM 範囲 " & _
'          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null and 終点側端末識別子 <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " GROUP BY 終点側端末識別子"

    j = 0
    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        
        If rs.RecordCount = 0 Then Stop
        
        Do Until rs.EOF
            Debug.Print rs(0), rs(1)
            If Not IsNull(rs(0)) Then
                ReDim Preserve サブ端末数RAN(1, j)
                For i = LBound(サブ端末数RAN, 1) To UBound(サブ端末数RAN, 1)
                    サブ端末数RAN(i, j) = rs(i)
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
Sub SQL_サブ確認_電線一覧_動作確認用temp(電線RAN, 製品品番str, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file 'ThisWorkbook.path & "\" & ThisWorkbook.Name
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    'Dim 範囲 As DataTable
    
    ReDim サブ端末数RAN(0, 0)
    Dim mysql(0) As String
    '始点側の回路

    mysql(0) = " SELECT [" & 製品品番str & "]" & _
          " FROM myTable" & _
          " WHERE [" & 製品品番str & "] Is Not Null" & _
          " ORDER BY [" & 製品品番str & "] ASC"
          
    'mySQL(0) = " SELECT " & Chr(34) & 製品品番str & Chr(34) & " ,COUNT(1)" & _
          " FROM 範囲" & _
          " WHERE 端末矢崎品番 is not null" '& _
          " GROUP BY " & Chr(34) & 製品品番str & Chr(34) & _
          " ORDER BY " & Chr(34) & 製品品番str & Chr(34) & " ASC"
    '終点側の回路
'    mySQL(1) = " SELECT 終点側端末識別子 , COUNT(1)" & _
'          " FROM 範囲 " & _
'          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null and 終点側端末識別子 <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" & _
'          " GROUP BY 終点側端末識別子"

    j = 0
    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        
        If rs.RecordCount = 0 Then Stop
        
        Do Until rs.EOF
            Debug.Print rs(0)
            If Not IsNull(rs(0)) Then
                ReDim Preserve サブ端末数RAN(0, j)
                If j = 45 Then Stop
                For i = LBound(サブ端末数RAN, 1) To UBound(サブ端末数RAN, 1)
                    サブ端末数RAN(i, j) = rs(i)
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

Sub SQL_サブ確認_電線一覧(電線RAN, 製品品番str, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With wb(0).Sheets("PVSW_RLTF")
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲2"
    End With
    
    
    ReDim 電線RAN(8, 0)
    Dim mysql(0) As String
    mysql(0) = " SELECT [" & 製品品番str & "],電線識別名 , 始点側端末矢崎品番 ,始点側端末識別子 , 終点側端末矢崎品番 ,終点側端末識別子 ,生区_,特区_,JCDF_" & _
          " FROM 範囲2 " & _
          " WHERE " & "[RLTFtoPVSW_]='Found'" & _
          " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
          " ORDER BY [" & 製品品番str & "] ASC"

    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        j = 0
        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            ReDim Preserve 電線RAN(8, j)
            For i = LBound(電線RAN, 1) To UBound(電線RAN, 1)
                電線RAN(i, j) = rs(i)
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_サブ図_先嵌め部品リスト_空栓(ran, ByVal 製品品番str, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    製品品番str = 製品品番str & String(15 - Len(製品品番str), " ")
    
    With wb(0).Sheets("PVSW_RLTF両端")
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim ran(1, 0)
    Dim mysql(0) As String
    mysql(0) = " SELECT [端末識別子],[EmptyPlug]" & _
          " FROM 範囲 " & _
          " WHERE [EmptyPlug] IS NOT NULL AND [EmptyPlug] <> """""

    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        j = 0
        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            j = j + 1
            ReDim Preserve ran(1, j)
            ran(0, j) = rs(0)
            ran(1, j) = rs(1)
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub


Sub SQL_製品別端末一覧(ran, 製品品番Ran, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(3, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For S = 1 To 製品品番RANc
        起動 = 製品品番Ran(製品品番RAN_read(製品品番Ran, "起動日"), S)
        '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
        If myTitle.Find(製品品番Ran(1, S), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & 製品品番Ran(1, S) & "],始点側端末矢崎品番 ,始点側端末識別子 ,'" & 製品品番Ran(1, S) & "'" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
        
            mysql(1) = " SELECT [" & 製品品番Ran(1, S) & "],終点側端末矢崎品番 ,終点側端末識別子 ,'" & 製品品番Ran(1, S) & "'" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
        
        
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                For r = LBound(ran, 2) To UBound(ran, 2)
                    If ran(0, r) = rs(1) Then
                        If ran(1, r) = rs(2) Then
                            If ran(2, r) = rs(3) Then
                                flg = True
                                Exit For
                            End If
                        End If
                    End If
                Next r
                '追加
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        If Not IsNull(rs(1)) Then
                            j = j + 1
                            ReDim Preserve ran(3, j)
                            ran(0, j) = rs(1)
                            ran(1, j) = rs(2)
                            ran(2, j) = rs(3)
                            ran(3, j) = 起動
                        End If
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next S
    cn.Close

End Sub

Sub SQL_電線一覧(ran, 製品品番Ran, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(7, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
    For S = 1 To 製品品番RANc
        '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
        If myTitle.Find(製品品番Ran(1, S), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 0
            mysql(0) = " SELECT [" & 製品品番Ran(1, S) & "],品種_,サイズ_,サ呼_,色_,色呼_,SA,'" & 製品品番Ran(1, S) & "'" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
        
'            mySQL(1) = " SELECT [" & 製品品番RAN(1, s) & "],終点側端末矢崎品番 ,終点側端末識別子 ,'" & 製品品番RAN(1, s) & "'" & _
'                  " FROM 範囲 " & _
'                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
'                  " AND [" & 製品品番RAN(1, s) & "] IS NOT NULL AND [" & 製品品番RAN(1, s) & "] <> """"" & _
'                  " ORDER BY [" & 製品品番RAN(1, s) & "] ASC"
        
        
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                '登録があるか確認
                For r = LBound(ran, 2) To UBound(ran, 2)
                    If ran(0, r) = rs(1) Then
                        If ran(1, r) = rs(2) Then
                            If ran(2, r) = rs(3) Then
                                If ran(3, r) = rs(4) Then
                                    If ran(4, r) = rs(5) Then
                                        If ran(5, r) = rs(6) Then
                                            If ran(6, r) = rs(7) Then
                                                flg = True
                                                ran(7, r) = ran(7, r) + 1
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next r
                '追加
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve ran(7, j)
                        ran(0, j) = rs(1)
                        ran(1, j) = rs(2)
                        ran(2, j) = rs(3)
                        ran(3, j) = rs(4)
                        ran(4, j) = rs(5)
                        ran(5, j) = rs(6)
                        ran(6, j) = rs(7) '製品品番
                        ran(7, j) = 1     '使用箇所数
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next S
    cn.Close

End Sub


Sub SQL_コネクタ一覧(ran, 製品品番Ran, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(4, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For S = 1 To 製品品番RANc
        '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
        If myTitle.Find(製品品番Ran(1, S), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & 製品品番Ran(1, S) & "],始点側端末矢崎品番,始点側端末識別子,TI1,'" & 製品品番Ran(1, S) & "'" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
            mysql(1) = " SELECT [" & 製品品番Ran(1, S) & "],終点側端末矢崎品番,終点側端末識別子,TI2,'" & 製品品番Ran(1, S) & "'" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
        
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                '登録があるか確認
                For r = LBound(ran, 2) To UBound(ran, 2)
                    If ran(0, r) = rs(1) Then
                        If ran(1, r) = rs(2) Then
                            If ran(2, r) = rs(3) Then
                                If ran(3, r) = rs(4) Then
                                    flg = True
                                    ran(4, r) = ran(4, r) + 1
                                End If
                            End If
                        End If
                    End If
                Next r
                '追加
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve ran(4, j)
                        ran(0, j) = rs(1)
                        ran(1, j) = rs(2)
                        ran(2, j) = rs(3)
                        ran(3, j) = rs(4)  '製品品番
                        ran(4, j) = 1      '使用箇所数
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next S
    cn.Close

End Sub

Sub SQL_挿入ガイド登録一覧(ran, 製品品番Ran, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(5, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For S = 1 To 製品品番RANc
        '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
        If myTitle.Find(製品品番Ran(1, S), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & 製品品番Ran(1, S) & "],始点側端末矢崎品番,始点側端末識別子,TI1,'" & 製品品番Ran(1, S) & "',TI_始点側挿入ガイド" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
            mysql(1) = " SELECT [" & 製品品番Ran(1, S) & "],終点側端末矢崎品番,終点側端末識別子,TI2,'" & 製品品番Ran(1, S) & "',TI_終点側挿入ガイド" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
        
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                '登録があるか確認
                For r = LBound(ran, 2) To UBound(ran, 2)
                    If ran(0, r) = rs(1) Then
                        If ran(1, r) = rs(2) Then
                            If ran(2, r) = rs(3) Then
                                If ran(3, r) = rs(4) Then
                                    If ran(5, r) = rs(5) Then
                                        flg = True
                                        ran(4, r) = ran(4, r) + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next r
                '追加
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve ran(5, j)
                        ran(0, j) = rs(1)
                        ran(1, j) = rs(2)
                        ran(2, j) = rs(3)
                        ran(3, j) = rs(4)  '製品品番
                        ran(4, j) = 1      '使用箇所数
                        ran(5, j) = rs(5)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next S
    cn.Close

End Sub
Sub SQL_YcEditor_Symbol(ran, myBook, 製品品番str)
    
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
    
    ReDim ran(3, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)

            mysql(0) = " SELECT 範囲b.[簡易ポイント],範囲a.[始点側回路符号],範囲a.[色_],範囲a.[色呼_]" & _
                  " FROM 範囲a INNER JOIN 範囲b" & _
                  " ON 範囲a.[始点側端末識別子] = 範囲b.[端末№] And 範囲a.[始点側端末矢崎品番] = 範囲b.[端末矢崎品番] AND 範囲a.[始点側キャビティ] = 範囲b.[Cav] " & _
                  " WHERE " & "範囲a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND 範囲a.[" & 製品品番str & "] IS NOT NULL AND 範囲a.[" & 製品品番str & "] <> """""
        
            mysql(0) = " SELECT 範囲b.簡易ポイント,範囲a.始点側回路符号,範囲a.色_,範囲a.色呼_" & _
                  " FROM 範囲a INNER JOIN 範囲b" & _
                  " ON 範囲a.始点側端末識別子 = 範囲b.端末№ And 範囲a.始点側端末矢崎品番 = 範囲b.端末矢崎品番 AND 範囲a.始点側キャビティ = 範囲b.Cav " & _
                  " WHERE " & "範囲a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND 範囲a.[" & 製品品番str & "] IS NOT NULL AND 範囲a.[" & 製品品番str & "] <> """""

                  
            mysql(1) = " SELECT 範囲b.簡易ポイント,範囲a.終点側回路符号,範囲a.色_,範囲a.色呼_" & _
                  " FROM 範囲a INNER JOIN 範囲b" & _
                  " ON 範囲a.終点側端末識別子 = 範囲b.端末№ And 範囲a.終点側端末矢崎品番 = 範囲b.端末矢崎品番 AND 範囲a.終点側キャビティ = 範囲b.Cav " & _
                  " WHERE " & "範囲a.[RLTFtoPVSW_] = 'Found'" & _
                  " AND 範囲a.[" & 製品品番str & "] IS NOT NULL AND 範囲a.[" & 製品品番str & "] <> """""
        For k = 0 To 1
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                '登録があるか確認
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
                '追加
                If flg = False Then
                    If rs(0) <> "" Then
                        j = j + 1
                        ReDim Preserve ran(3, j)
                        ran(0, j) = rs(0)
                        ran(1, j) = rs(1)
                        ran(2, j) = rs(2)
                        ran(3, j) = rs(3)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
    cn.Close

End Sub

Sub SQL_YcEditor_WH(ran, myBook, 製品品番str)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim ran(4, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
        '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
        For k = 0 To 0
            mysql(0) = " SELECT 構成_,始点側回路符号,終点側回路符号,色_,色呼_" & _
                  " FROM 範囲" & _
                  " WHERE " & "[RLTFtoPVSW_] = 'Found'" & _
                  " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """""
                  
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                '登録があるか確認
                For r = LBound(ran, 2) To UBound(ran, 2)
                    If ran(0, r) = rs(0) Then
                        If ran(1, r) = rs(1) Then
                            If ran(2, r) = rs(2) Then
                                If ran(3, r) = rs(3) Then
                                    If ran(4, r) = rs(4) Then
                                        flg = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next r
                '追加
                If flg = False Then
                    If rs(0) <> "" Then
                        j = j + 1
                        ReDim Preserve ran(4, j)
                        ran(0, j) = rs(0)
                        ran(1, j) = rs(1)
                        ran(2, j) = rs(2)
                        ran(3, j) = rs(3)
                        ran(4, j) = rs(4)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
    cn.Close

End Sub



Sub SQL_挿入ガイド一覧(ran, 製品品番Ran, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(5, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For S = 1 To 製品品番RANc
        '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
        If myTitle.Find(製品品番Ran(1, S), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & 製品品番Ran(1, S) & "],始点側端末矢崎品番,始点側端末識別子,TI1,'" & 製品品番Ran(1, S) & "',TI_始点側挿入ガイド" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
            mysql(1) = " SELECT [" & 製品品番Ran(1, S) & "],終点側端末矢崎品番,終点側端末識別子,TI2,'" & 製品品番Ran(1, S) & "',TI_終点側挿入ガイド" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
        
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                '登録があるか確認
                For r = LBound(ran, 2) To UBound(ran, 2)
'                    If ran(0, r) = rs(1) Then
'                        If ran(1, r) = rs(2) Then
                            If ran(2, r) = rs(3) Then
                                If ran(3, r) = rs(4) Then
                                    If ran(5, r) = rs(5) Then
                                        flg = True
                                        ran(4, r) = ran(4, r) + 1
                                    End If
                                End If
                            End If
'                        End If
'                    End If
                Next r
                '追加
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve ran(5, j)
                        ran(0, j) = rs(1)
                        ran(1, j) = rs(2)
                        ran(2, j) = rs(3)
                        ran(3, j) = rs(4)  '製品品番
                        ran(4, j) = 1      '使用箇所数
                        ran(5, j) = rs(5)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next S
    cn.Close

End Sub

Sub SQL_端子一覧(ran, 製品品番Ran, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(5, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For S = 1 To 製品品番RANc
        '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
        If myTitle.Find(製品品番Ran(1, S), , , 1) Is Nothing Then GoTo nexts
        For k = 0 To 1
            mysql(0) = " SELECT [" & 製品品番Ran(1, S) & "],始点側端子_,始点側部品_,始点側メ_,SM1,'" & 製品品番Ran(1, S) & "'" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
        
             mysql(1) = " SELECT [" & 製品品番Ran(1, S) & "],終点側端子_,終点側部品_,終点側メ_,SM2,'" & 製品品番Ran(1, S) & "'" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
                  " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
        
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                '登録があるか確認
                For r = LBound(ran, 2) To UBound(ran, 2)
                    If ran(0, r) = rs(1) Then
                        If ran(1, r) = rs(2) Then
                            If ran(2, r) = rs(3) Then
                                If ran(3, r) = rs(4) Then
                                    If ran(4, r) = rs(5) Then
                                        flg = True
                                        ran(5, r) = ran(5, r) + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next r
                '追加
                If flg = False Then
                    If rs(1) & rs(2) <> "" Then
                        j = j + 1
                        ReDim Preserve ran(5, j)
                        ran(0, j) = rs(1)
                        ran(1, j) = rs(2)
                        ran(2, j) = rs(3)
                        ran(3, j) = rs(4)
                        ran(4, j) = rs(5) '製品品番
                        ran(5, j) = 1     '使用箇所数
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next S
    cn.Close

End Sub
Sub SQL_端末サブ一覧(ran, 製品品番str, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(2, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
    If myTitle.Find(製品品番str, , , 1) Is Nothing Then GoTo nexts
    For k = 0 To 1
        mysql(0) = " SELECT [" & 製品品番str & "],始点側端末識別子,始点側端末矢崎品番" & _
              " FROM 範囲 " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
              " ORDER BY [" & 製品品番str & "] ASC"
    
        mysql(1) = " SELECT [" & 製品品番str & "],終点側端末識別子,終点側端末矢崎品番" & _
              " FROM 範囲 " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
              " ORDER BY [" & 製品品番str & "] ASC"
    
        'SQLを開く
        rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            '登録があるか確認
            If Mid(rs(2), 1, 4) <> "7009" Then GoTo line10
            '追加
                If rs(1) & rs(2) <> "" Then
                    j = j + 1
                    ReDim Preserve ran(2, j)
                    ran(0, j) = rs(1)
                    ran(1, j) = rs(2)
                    ran(2, j) = rs(0)
            End If
line10:
            rs.MoveNext
        Loop
        rs.Close
    Next k
nexts:
    cn.Close
End Sub
Sub SQL_製品別端末一覧_防水(ran, 製品品番Ran, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(2, 0): j = 0
    Dim mysql() As String: ReDim mysql(1)
    For S = 1 To 製品品番RANc
        '[製品品番]から見て[PVSW_RLTF]にメイン品番が無い時、処理を飛ばす
        If myTitle.Find(製品品番Ran(1, S), , , 1) Is Nothing Then GoTo nexts
        mysql(0) = " SELECT [" & 製品品番Ran(1, S) & "],[始点側端末矢崎品番],[始点側端末識別子] " & _
              " FROM 範囲 " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
              " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
    
        mysql(1) = " SELECT [" & 製品品番Ran(1, S) & "],[終点側端末矢崎品番],[終点側端末識別子] " & _
              " FROM 範囲 " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & 製品品番Ran(1, S) & "] IS NOT NULL AND [" & 製品品番Ran(1, S) & "] <> """"" & _
              " ORDER BY [" & 製品品番Ran(1, S) & "] ASC"
        For k = 0 To 1
        
            'SQLを開く
            rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
            If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                           'PVSW_RLTFの書式設定を@にするとか
            Do Until rs.EOF
                flg = False
                For r = LBound(ran, 2) To UBound(ran, 2)
                    If ran(0, r) = rs(1) And ran(1, r) = rs(2) Then
                        flg = True
                        Exit For
                    End If
                Next r
                '追加
                If flg = False Then
                    If rs(1) <> "" Then
                        j = j + 1
                        ReDim Preserve ran(2, j)
                        ran(0, j) = rs(1)
                        ran(1, j) = rs(2)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next k
nexts:
    Next S
    cn.Close

End Sub


Sub SQL_csvインポート(対象ファイル, myBookpath)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Text;HRD=YES;FMT=Delimited"
    cn.Open Left(myBookpath, InStrRev(myBookpath, "\")) & "000_システムパーツ\"
    Set rs = New ADODB.Recordset
    
    ReDim 電線RAN(5, 0)
    Dim mysql(0) As String
    mysql(0) = " SELECT * " & _
          " FROM " & 対象ファイル '& _
          " WHERE " & "[種類]='写真'" ' & _
          " AND [" & 製品品番str & "] IS NOT NULL" & _
          " ORDER BY [" & 製品品番str & "] ASC"

    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        'ワークシートの追加
        For Each ws(0) In Worksheets
            If ws(0).Name = 対象ファイル Then
                Application.DisplayAlerts = False
                ws(0).Delete
                Application.DisplayAlerts = True
            End If
        Next ws
        Set newSheet = Worksheets.add
        newSheet.Name = 対象ファイル
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
'                                       'PVSW_RLTFの書式設定を@にするとか
'        Do Until rs.EOF
'            ReDim Preserve 電線RAN(5, J)
'            For i = LBound(電線RAN, 1) To UBound(電線RAN, 1)
'                電線RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop
        With newSheet
            .Cells.NumberFormat = "@"
            For i = 0 To rs.fields.count - 1
                .Cells(1, i + 1) = rs(i).Name
            Next i
            .Range("a2").CopyFromRecordset rs
        End With
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_製品別端末一覧_CAV座標(ran, 部品品番str, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Text;HRD=YES:FMT=Delimited"
    cn.Open myAddress(1, 1) & "\"
    Set rs = New ADODB.Recordset
    
    ReDim ran(5, 0): j = 0
    Dim mysql(1) As String
    mysql(0) = " SELECT [PartName],[Cav],[Width],[Height],[EmptyPlug],[PlugColor] " & _
          " FROM CAV座標.txt" & _
          " WHERE [PartName]='" & 部品品番str & "'" & _
             "AND [種類]='写真'" & _
          " ORDER BY [Cav] ASC" ' & _
          " GROUP BY [Cav]"
    
    mysql(1) = " SELECT [PartName],[Cav],[Width],[Height],[EmptyPlug],[PlugColor] " & _
          " FROM CAV座標.txt" & _
          " WHERE [PartName]='" & 部品品番str & "'" & _
             "AND [種類]='略図'" & _
          " ORDER BY [CAV] ASC" '& _
          " GROUP BY [PartName],[Cav]"
    For a = 0 To 1
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
        If rs.RecordCount = 0 Then GoTo line20
        Do Until rs.EOF
            j = j + 1
            ReDim Preserve ran(5, j)
            ran(0, j) = rs(0)
            ran(1, j) = rs(1)
            ran(2, j) = rs(2)
            ran(3, j) = rs(3)
            ran(4, j) = rs(4)
            ran(5, j) = rs(5)
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

Public Function SQL_製品別端末一覧_CAV座標2(ran, 部品品番str, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Text;HRD=YES:FMT=Delimited"
    
    cn.Open myAddress(1, 1) & "\200_CAV座標\"
    Set rs = New ADODB.Recordset
    
    ReDim ran(5, 0): j = 0
    Dim mysql(1) As String
    mysql(0) = " SELECT [PartName],[Cav],[Width],[Height],[EmptyPlug],[PlugColor] " & _
          " FROM " & "'" & 部品品番str & "'" '& _
          " WHERE [PartName]='" & 部品品番str & "'" & _
          " ORDER BY [Cav] ASC" ' & _
          " GROUP BY [Cav]"
    
    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
        If rs.RecordCount = 0 Then GoTo line20
        Do Until rs.EOF
            j = j + 1
            ReDim Preserve ran(5, j)
            ran(0, j) = rs(0)
            ran(1, j) = rs(1)
            ran(2, j) = rs(2)
            ran(3, j) = rs(3)
            ran(4, j) = rs(4)
            ran(5, j) = rs(5)
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
    SQL_製品別端末一覧_CAV座標2 = j
End Function

Sub SQL_サブナンバー印刷_データ作成(製品品番Ran, mySheet, tempアドレス, ByVal myPosSP As Variant, ByVal kumitateList As Variant, Optional ByVal CheckBox_stepNumberAdd As Boolean)
    
    If CheckBox_stepNumberAdd Then
        Dim myRan As Variant, myPath As String, 設変str As String
        myPath = wb(0).path & dirString_09 & Replace(製品品番str, " ", "") & "_wire.txt"
        myRan = readTextToArray(myPath)
        'myRan = WorksheetFunction.Transpose(myRan)
    End If
    
     '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long
        lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim ran(2, 0)
    j = 0
    Dim mysql(0) As String, myCount As Long
    
    For r = LBound(製品品番Ran, 2) + 1 To UBound(製品品番Ran, 2)
        サブ印刷 = 製品品番Ran(製品品番RAN_read(製品品番Ran, "サブ"), r)
        品番str = 製品品番Ran(製品品番RAN_read(製品品番Ran, "メイン品番"), r)
        If サブ印刷 = "1" Or 品番str = 製品品番str Then
            mysql(0) = " SELECT [" & 品番str & "],left(電線識別名,4),'" & Replace(品番str, " ", "") & "'" & _
                  " FROM 範囲 " & _
                  " WHERE " & "[RLTFtoPVSW_]='Found'" & _
                  " AND [" & 品番str & "] IS NOT NULL AND [" & 品番str & "] <> """"" & _
                  " ORDER BY [電線識別名] ASC"
            For a = 0 To 0
                'SQLを開く
                rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
                If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                                                   'PVSW_RLTFの書式設定を@にするとか
                Do Until rs.EOF
                    ReDim Preserve ran(2, j)
                    For i = LBound(ran, 1) To UBound(ran, 1)
                        ran(i, j) = rs(i)
                    Next i
                    j = j + 1
                    rs.MoveNext
                Loop
                rs.Close
            Next a
            myCount = myCount + 1
        End If
    Next r
    cn.Close
    
    If myCount = 0 Then Stop ' 有効な製品品番が無い
    
    'テキストファイルにして出力
    Dim lntFlNo As Integer: lntFlNo = FreeFile
    Dim outPutAddress As String: outPutAddress = tempアドレス
    Open outPutAddress For Output As #lntFlNo
    Dim myLine As Variant, subSubNumber As String
    日時 = now
    For i = LBound(ran, 2) To UBound(ran, 2)
        構成 = ran(1, i)
        If CheckBox_stepNumberAdd Then
            subSubNumber = searchRan_ver2(myRan, 構成, "構成_", "subSubNumber")
            If subSubNumber = "False" Or ran(0, i) = "999" Then
                サブ値 = ran(0, i)
            Else
                サブ値 = ran(0, i) & "-" & subSubNumber
            End If
        Else
            サブ値 = ran(0, i)
        End If
        製品 = ran(2, i)
        For r = LBound(kumitateList, 2) + 1 To UBound(kumitateList, 2)
            myLine = Empty
            For ii = LBound(myPosSP) To UBound(myPosSP)
                If myPosSP(ii) <> "" Then
                    Select Case ii
                        Case 1
                            myVal = 製品
                        Case 3
                            myVal = 構成
                        Case 4
                            myVal = サブ値
                        Case 5
                            myVal = kumitateList(0, r)
                        Case Else
                            myVal = "" '切断と設変_0と2
                    End Select
                    myLine = myLine & myVal & Chr(44)
                End If
            Next ii
            myLine = myLine & 日時
            Print #lntFlNo, myLine
        Next r
    Next i
    
    Close lntFlNo

End Sub

Sub SQL_変更依頼_線長(製品品番Ran, 線長変更RAN, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    Dim mysql(0) As String
    
    mysql(0) = "SELECT "
    For i = 1 To 製品品番RANc
        mysql(0) = mysql(0) & "[" & 製品品番Ran(1, i - 1) & "],"
    Next i
    
    ReDim 線長変更RAN(製品品番RANc + 6, 0)
    mysql(0) = mysql(0) & "構成_,始点側回路符号, 終点側回路符号, 線長_ ,線長後_ ,RLTFtoPVSW_,備考_" & _
          " FROM myTable " & _
          " WHERE " & "[RLTFtoPVSW_]='Found'" & _
          " AND [線長後_] IS NOT NULL"

    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        j = 0
        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            ReDim Preserve 線長変更RAN(製品品番RANc + 6, j)
            For i = LBound(線長変更RAN, 1) To UBound(線長変更RAN, 1)
                線長変更RAN(i, j) = rs(i)
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_端末一覧_2(製品品番Ran, 電線一覧RAN, 端末一覧ran, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    Dim mysql(0) As String
    
    mysql(0) = "SELECT "
    For i = 1 To 製品品番RANc
        mysql(0) = mysql(0) & "[" & 製品品番Ran(1, i - 1) & "],"
    Next i
    
    ReDim 電線一覧RAN(製品品番RANc + 9, 0)
    ReDim 端末一覧ran(0)
    mysql(0) = mysql(0) & "構成_,始点側回路符号, 終点側回路符号, 始点側端末識別子, 終点側端末識別子,始点側キャビティ,終点側キャビティ,線長_,線長後_ ,RLTFtoPVSW_,備考_" & _
          " FROM myTable " & _
          " WHERE " & "[RLTFtoPVSW_]='Found'" '& _
          " AND [線長後_] IS NOT NULL"

    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim j As Long: j = 0
        Dim jj As Long: jj = 0
        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            '冶具typeの対象にあるか確認
            findFlg = False
            For i = 1 To 製品品番RANc
                If Not IsNull(rs(i - 1)) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            
            If findFlg = False Then
                GoTo line20
            End If
            
            '追加
            ReDim Preserve 電線一覧RAN(製品品番RANc + 9, j)
            
            For i = 1 To 製品品番RANc
                電線一覧RAN(i - 1, j + 0) = rs(i - 1)
            Next i
                '始点
                電線一覧RAN(製品品番RANc + 0, j + 0) = rs(製品品番RANc + 0) '構成
                電線一覧RAN(製品品番RANc + 1, j + 0) = rs(製品品番RANc + 1) '回符
                電線一覧RAN(製品品番RANc + 2, j + 0) = rs(製品品番RANc + 2)
                電線一覧RAN(製品品番RANc + 3, j + 0) = rs(製品品番RANc + 3) '端末
                電線一覧RAN(製品品番RANc + 4, j + 0) = rs(製品品番RANc + 4)
                電線一覧RAN(製品品番RANc + 5, j + 0) = rs(製品品番RANc + 5) 'cav
                電線一覧RAN(製品品番RANc + 6, j + 0) = rs(製品品番RANc + 6)
                電線一覧RAN(製品品番RANc + 7, j + 0) = rs(製品品番RANc + 7) '線長_
                電線一覧RAN(製品品番RANc + 8, j + 0) = rs(製品品番RANc + 8) '線長後_
                電線一覧RAN(製品品番RANc + 9, j + 0) = rs(製品品番RANc + 10) '備考_
                
            '始点端末無い時追加
            For i = LBound(端末一覧ran) To UBound(端末一覧ran)
                findFlg = False
                If 端末一覧ran(i) = rs(製品品番RANc + 3) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            If findFlg = False Then
                ReDim Preserve 端末一覧ran(jj)
                端末一覧ran(jj) = rs(製品品番RANc + 3)
                jj = jj + 1
            End If
            '終点端末無い時追加
            For i = LBound(端末一覧ran) To UBound(端末一覧ran)
                findFlg = False
                If 端末一覧ran(i) = rs(製品品番RANc + 4) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            If findFlg = False Then
                ReDim Preserve 端末一覧ran(jj)
                端末一覧ran(jj) = rs(製品品番RANc + 4)
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
Sub SQL_ハメ図作成_1(製品品番Ran, ハメ図作成RAN, 端末, myBook, newSheet)
    
    Call SQL_csvインポート("部材詳細.txt", myBook.path)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open ThisWorkbook.FullName
    Set rs = New ADODB.Recordset
    
    'myTable0
    With myBook.Sheets("PVSW_RLTF")
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable0"
    End With
    
    'myTable1
    With myBook.Sheets("ポイント一覧")
        Set key = .Cells.Find("端末矢崎品番", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable1"
    End With
    
    'myTable2
    With myBook.Sheets("製品別端末一覧")
        Set key = .Cells.Find("防水コネクタ品番", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(firstRow, key.Column), .Cells(lastRow, lastCol)).Name = "myTable2"
        Set key = Nothing
    End With
    
    'myTable3
    With myBook.Sheets("部材詳細.txt")
        Set key = .Cells.Find("部品品番_", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(firstRow, key.Column), .Cells(lastRow, lastCol)).Name = "myTable3"
        Set key = Nothing
    End With
    
    Dim 電線一覧RAN() As String
    ReDim 電線一覧RAN(製品品番RANc + 11, 0)
    Dim mysql(1) As String
    
    For a = LBound(mysql) To UBound(mysql)
        mysql(a) = "SELECT "
        For i = 1 To 製品品番RANc
            mysql(a) = mysql(a) & "[" & 製品品番Ran(1, i - 1) & "],"
        Next i
    Next a
    
    mysql(0) = mysql(0) & "構成_,始点側回路符号, 始点側端末識別子,始点側キャビティ,始点側端末矢崎品番,線長_,線長後_ ,RLTFtoPVSW_,始点側マ_,色呼_,品種_,サ呼_,両端ハメ,両端同端子,始点側ハメ,始点側相手_,始点側同_,'始' AS 側" & _
                          ",b.[ポイント1]" & _
                          ",c.[EmptyPlug],c.[PlugColor]" & _
                          ",d.[コネクタ極数_]" & _
          " FROM (((myTable0 AS a" & _
          " LEFT OUTER JOIN myTable1 AS b " & _
          " ON a.[始点側端末矢崎品番] = b.[端末矢崎品番] AND a.[始点側端末識別子] = b.[端末№] AND a.[始点側キャビティ] = b.[Cav])" & _
          " LEFT OUTER JOIN myTable2 AS c " & _
          " ON a.[始点側端末矢崎品番] = c.[防水コネクタ品番] AND a.[始点側端末識別子] = c.[端末№_] AND a.[始点側キャビティ] = c.[Cav])" & _
          " LEFT OUTER JOIN myTable3 AS d " & _
          " ON a.[始点側端末矢崎品番] = d.[部品品番_] )" & _
          " WHERE " & "a.[RLTFtoPVSW_]='Found' AND a.[始点側端末識別子] is not Null AND a.[始点側キャビティ] is not Null"

    mysql(1) = mysql(1) & "構成_,終点側回路符号, 終点側端末識別子,終点側キャビティ,終点側端末矢崎品番,線長_,線長後_ ,RLTFtoPVSW_,終点側マ_,色呼_,品種_,サ呼_,両端ハメ,両端同端子,終点側ハメ,終点側相手_,終点側同_,'終' AS 側" & _
                          ",b.[ポイント1]" & _
                          ",c.[EmptyPlug],c.[PlugColor]" & _
                          ",d.[コネクタ極数_]" & _
          " FROM (((myTable0 AS a" & _
          " LEFT OUTER JOIN myTable1 AS b " & _
          " ON a.[終点側端末矢崎品番] = b.[端末矢崎品番] AND a.[終点側端末識別子] = b.[端末№] AND a.[終点側キャビティ] = b.[Cav])" & _
          " LEFT OUTER JOIN myTable2 AS c " & _
          " ON a.[終点側端末矢崎品番] = c.[防水コネクタ品番] AND a.[終点側端末識別子] = c.[端末№_] AND a.[終点側キャビティ] = c.[Cav])" & _
          " LEFT OUTER JOIN myTable3 AS d " & _
          " ON a.[終点側端末矢崎品番] = d.[部品品番_] )" & _
          " WHERE " & "a.[RLTFtoPVSW_]='Found' AND a.[終点側端末識別子] is not Null AND a.[終点側キャビティ] is not Null"

    For a = LBound(mysql) To UBound(mysql)
        For i = 1 To 製品品番RANc
            If i = 1 Then
                mysql(a) = mysql(a) & " AND [" & 製品品番Ran(1, i - 1) & "] is not null"
            Else
                mysql(a) = mysql(a) & " OR [" & 製品品番Ran(1, i - 1) & "] is not null"
            End If
        Next i
    Next a
          
    'mySQL(0) = mySQL(0) & " ORDER BY [始点側端末識別子] ASC , [始点側キャビティ] ASC"

    For a = LBound(mysql) To UBound(mysql)
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim j As Long: j = 0
        Dim jj As Long: jj = 0
        'ワークシートの追加
        If a = LBound(mysql) Then
            For Each ws(0) In Worksheets
                If ws(0).Name = "ハメ図temp" Then
                    Application.DisplayAlerts = False
                    ws(0).Delete
                    Application.DisplayAlerts = True
                End If
            Next ws
            Set newSheet = Worksheets.add
            newSheet.Name = "ハメ図temp"
        End If
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
'                                       'PVSW_RLTFの書式設定を@にするとか
'        Do Until rs.EOF
'            ReDim Preserve 電線RAN(5, J)
'            For i = LBound(電線RAN, 1) To UBound(電線RAN, 1)
'                電線RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop
        With newSheet
            .Cells.NumberFormat = "@"
            For i = 0 To rs.fields.count - 1
                .Cells(1, i + 1) = Replace(Replace(rs(i).Name, "始点側", ""), "終点側", "")
            Next i
            lastRow = .Cells(.Rows.count, .Cells.Find("構成_", , , 1).Column).End(xlUp).Row + 1
            .Cells(lastRow, 1).CopyFromRecordset rs
        End With
        Debug.Print rs.RecordCount
        rs.Close
    Next a
    cn.Close
    
End Sub

Sub SQL_ハメ図作成_2(製品品番Ran, myBook, newSheet)
    
    Call SQL_csvインポート("CAV座標.txt", myBook.path)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
    With myBook.Sheets("CAV座標.txt")
        Set key = .Cells.Find("PartName", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable1"
    End With
    
    Dim mysql(0) As String
    mysql(0) = "SELECT a.*,b.[x],b.[種類]" & _
          " FROM myTable1 AS a" & _
          " LEFT OUTER JOIN myTable0 AS b " & _
          " ON a.[端末矢崎品番] = b.[PartName] AND a.[キャビティ] = b.[Cav] " & _
          " WHERE b.[種類] = '写真'" 'a.[RLTFtoPVSW_]='Found' AND a.[始点側端末識別子] is not Null AND a.[始点側キャビティ] is not Null"
          
    'mySQL(1) = "SELECT a.* " & _
                     ",b.[x] ,b.[種類]" & _
          " FROM myTable0 AS a" & _
          " LEFT OUTER JOIN myTable1 AS b " & _
          " ON a.[端末矢崎品番] = b.[PartName] AND a.[キャビティ] = b.[Cav]" & _
          " WHERE b.[種類] = '略図'" 'a.[RLTFtoPVSW_]='Found' AND a.[始点側端末識別子] is not Null AND a.[始点側キャビティ] is not Null"
          
    'mySQL(0) = mySQL(0) & " ORDER BY [始点側端末識別子] ASC , [始点側キャビティ] ASC"

    For a = LBound(mysql) To UBound(mysql)
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim j As Long: j = 0
        Dim jj As Long: jj = 0
        'セルの値を削除
        If a = LBound(mysql) Then
            'ワークシートの追加
            If a = LBound(mysql) Then
                For Each ws(0) In Worksheets
                    If ws(0).Name = "ハメ図temp1" Then
                        Application.DisplayAlerts = False
                        ws(0).Delete
                        Application.DisplayAlerts = True
                    End If
                Next ws
                Set newSheet = Worksheets.add
                newSheet.Name = "ハメ図temp1"
            End If
        End If
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
'                                       'PVSW_RLTFの書式設定を@にするとか
'        Do Until rs.EOF
'            ReDim Preserve 電線RAN(5, J)
'            For i = LBound(電線RAN, 1) To UBound(電線RAN, 1)
'                電線RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop

'        With newSheet
'            .Cells.NumberFormat = "@"
'            For i = 0 To rs.Fields.count - 1
'                .Cells(1, i + 1) = Replace(Replace(rs(i).Name, "始点側", ""), "終点側", "")
'            Next i
'            lastRow = .UsedRange.Rows.count + 1
'            Debug.Print rs.RecordCount
'            .Cells(lastRow, 1).CopyFromRecordset rs
'        End With
        rs.Close
    Next a
    cn.Close
    
End Sub

Sub SQL_サブナンバー印刷_データ更新(tempアドレス, tempアドレス2, tempアドレス3, ByVal mySqlOn)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    'ヘッダーの無いテキストファイルの時 12.0だとフィールド名がF1でとれない
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=NO;FMT=Delimited"
    cn.Open Left(tempアドレス, InStrRev(tempアドレス, "\") - 1)
    Set rs = New ADODB.Recordset

    Dim mysql(2) As String
    'change(1に含むから不要になった)
    mysql(0) = " SELECT b.* " & _
          " FROM " & Mid(tempアドレス2, InStrRev(tempアドレス2, "\") + 1) & " as b" & _
          " INNER JOIN " & Mid(tempアドレス, InStrRev(tempアドレス, "\") + 1) & " as a" & _
          " ON a.F2 = b.F2 AND a.F4 = b.F4 "
    'newとchange
    mysql(1) = " SELECT b.* " & _
          " FROM " & Mid(tempアドレス2, InStrRev(tempアドレス2, "\") + 1) & " as b" & _
          " LEFT OUTER JOIN " & Mid(tempアドレス, InStrRev(tempアドレス, "\") + 1) & " as a" & _
          mySqlOn(0)
    'old
    mysql(2) = " SELECT a.* " & _
          " FROM " & Mid(tempアドレス2, InStrRev(tempアドレス2, "\") + 1) & " as b" & _
          " RIGHT OUTER JOIN " & Mid(tempアドレス, InStrRev(tempアドレス, "\") + 1) & " as a" & _
          mySqlOn(1)
    
    Dim サブ印刷Ran() As Variant
    For a = 1 To UBound(mysql)
        'SQLを開く
        'cn.Execute mySQL(0)
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        If a = 1 Then ReDim サブ印刷Ran(rs.fields.count - 1, 0): j = 0
        'Sheets("Sheet1").Cells.ClearContents
        Do Until rs.EOF
            ReDim Preserve サブ印刷Ran(rs.fields.count - 1, j)
            For i = 0 To rs.fields.count - 1
                'Sheets("Sheet1").Cells(J + 1, i + 1) = rs(i).Value
                サブ印刷Ran(i, j) = rs(i).Value
            Next i
            j = j + 1
            'Range("a2").CopyFromRecordset rs
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
    'ファイル作成
    Dim lntFlNo As Integer: lntFlNo = FreeFile
    Open tempアドレス3 For Output As #lntFlNo
    
    Dim サブ値 As String, 構成 As String, 製品 As String
    Dim 日時 As Date
    Dim x As Long, y As Long, fndX As Long
    
    For x = LBound(サブ印刷Ran, 2) To UBound(サブ印刷Ran, 2)
        If Not IsNull(サブ印刷Ran(1, x)) Then
        myLine = Empty
        For xx = LBound(サブ印刷Ran) To UBound(サブ印刷Ran)
            If xx <> UBound(サブ印刷Ran) Then
                myLine = myLine & サブ印刷Ran(xx, x) & Chr(44)
            Else
                myLine = myLine & サブ印刷Ran(xx, x) '最後は日時
            End If
        Next xx
        Print #lntFlNo, myLine
line20:
        End If
    Next x
    
    Close #lntFlNo
    
End Sub
Public Function SQL_MDファイル読み込み_空栓(製品品番str, 設変str, myRan)
    製品品番str = Replace(製品品番str, " ", "")
    tempアドレス1 = ThisWorkbook.path & "\08_MD\" & 製品品番str & "_" & 設変str & "_MD" & "\004Term.csv"
    tempアドレス2 = ThisWorkbook.path & "\08_MD\" & 製品品番str & "_" & 設変str & "_MD" & "\006Cone.csv"
    If Dir(tempアドレス1) = "" Then Exit Function
    If Dir(tempアドレス2) = "" Then Exit Function
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=YES;FMT=Delimited"
    cn.Open Left(tempアドレス1, InStrRev(tempアドレス1, "\") - 1)
    Set rs = New ADODB.Recordset

    Dim mysql(0) As String
    mysql(0) = " SELECT a.部品品番,a.サブ番号,a.キャビティ番号,a.投入工程,b.コネクタ番号 ,b.部品品番" & _
          " FROM " & Mid(tempアドレス1, InStrRev(tempアドレス1, "\") + 1) & " as a" & _
          " INNER JOIN " & Mid(tempアドレス2, InStrRev(tempアドレス2, "\") + 1) & " as b" & _
          " ON a.取付け先ＩＤ = b.ＩＤ " 'AND a.F4 = b.F4 "
    j = 0
    For a = 0 To UBound(mysql)
        'SQLを開く
        'cn.Execute mySQL(0)
        On Error Resume Next
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        If err.number = -2147467259 Then GoTo line20
        On Error GoTo 0
        If a = 0 Then ReDim myRan(rs.fields.count, 0): j = 0
        
        Do Until rs.EOF
            ReDim Preserve myRan(rs.fields.count, j)
            For i = 0 To rs.fields.count - 1
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
    
    SQL_MDファイル読み込み_空栓 = UBound(myRan, 2)
End Function

Sub SQL_test()
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    'ヘッダーの無いテキストファイルの時 12.0だとフィールド名がF1でとれない
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=NO;FMT=Delimited"
    cn.Open "D:\04_製品の動き\028_675W_543B"
    Set rs = New ADODB.Recordset
    
    '同条件の時、更新
    ReDim 電線RAN(5, 0)
    Dim mysql(0) As String
    mysql(0) = " SELECT * " & _
          " FROM efu_subNo_temp3.txt " & _
          " WHERE F6 in " & _
          " ( SELECT MAX(F6) FROM efu_subNo_temp3.txt GROUP BY F2,F4 ORDER BY F2,F4)" '& _
          " INNER JOIN " & Mid(tempアドレス, InStrRev(tempアドレス, "\") + 1) & " as b" & _
          " ON a.F2=b.F2 AND a.F4 = b.F4" '& _
          " SET a.F4 = b.F4" & _
          " WHERE a.F2=b.F2 AND a.F4 = b.F4" ' & _
          " AND [" & 製品品番str & "] IS NOT NULL" & _
          " ORDER BY [" & 製品品番str & "] ASC"
          'mySQL(0) = "SELECT MAX(F6),F2,F4 FROM efu_subNo_temp3.txt GROUP BY F2,F4 ORDER BY F2,F4"

    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        'ワークシートの追加
        For Each ws(0) In Worksheets
            If ws(0).Name = 対象ファイル Then
                Application.DisplayAlerts = False
                ws(0).Delete
                Application.DisplayAlerts = True
            End If
        Next ws
        Set newSheet = Worksheets.add
        newSheet.Name = 対象ファイル
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
'                                       'PVSW_RLTFの書式設定を@にするとか
'        Do Until rs.EOF
'            ReDim Preserve 電線RAN(5, J)
'            For i = LBound(電線RAN, 1) To UBound(電線RAN, 1)
'                電線RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop
        With newSheet
            .Cells.NumberFormat = "@"
            For i = 0 To rs.fields.count - 1
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
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\04_製品の動き\028_675W_543B;" _
                                & "Extended Properties='text;HDR=No;FMT=Delimited'"
            .Open
        End With
    
    Set rec = CreateObject("ADODB.Recordset")
        rec.Open "select * from efu_subNo_temp2.txt as a where a.[F2] ='821113B300'", con
        Debug.Print rec(1) '最初のレコードの1列目の値を表示

End Sub

Sub SQL_CAV座標取得(製品品番Ran, myBook, newSheet)
    
    Call SQL_csvインポート("CAV座標.txt", myBook.path)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
    With myBook.Sheets("CAV座標.txt")
        Set key = .Cells.Find("PartName", , , 1)
        firstRow = key.Row
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable0"
    End With
    
    Dim mysql(0) As String
    mysql(0) = "SELECT a.[PartName],a.[Cav],a.[X],a.[Y],a.[Width],a.[Height],a.[形状],a.[種類],a.[Angle],a.[Width(mm)],a.[Category],a.[Rock]" & _
          " FROM myTable0 AS a" & _
          " LEFT JOIN myTable1 AS b " & _
          " ON a.[PartName] = b.[端末矢崎品番] AND a.[Cav] = b.[キャビティ]" & _
          " WHERE a.[PartName] is not Null" 'a.[RLTFtoPVSW_]='Found' AND a.[始点側端末識別子] is not Null AND a.[始点側キャビティ] is not Null"
          
    'mySQL(1) = "SELECT a.* " & _
                     ",b.[x] ,b.[種類]" & _
          " FROM myTable0 AS a" & _
          " LEFT OUTER JOIN myTable1 AS b " & _
          " ON a.[端末矢崎品番] = b.[PartName] AND a.[キャビティ] = b.[Cav]" & _
          " WHERE b.[種類] = '略図'" 'a.[RLTFtoPVSW_]='Found' AND a.[始点側端末識別子] is not Null AND a.[始点側キャビティ] is not Null"
          
    'mySQL(0) = mySQL(0) & " ORDER BY [始点側端末識別子] ASC , [始点側キャビティ] ASC"

    For a = LBound(mysql) To UBound(mysql)
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim addCol() As Long, 追加F
        Dim cav As String
        With newSheet
            追加F = "X,Y,Width,Height,形状,種類,Angle,Width(mm),Category,Rock"
            ReDim addCol(rs.fields.count - 1)

            For x = 1 To rs.fields.count
                If InStr(追加F, rs(x - 1).Name) > 0 Then
                    addCol(x - 1) = .Cells(1, .Columns.count).End(xlToLeft).Column + 1
                    .Cells(1, addCol(x - 1)) = rs(x - 1).Name
                Else
                    addCol(x - 1) = 0
                End If
            Next x
            矢崎Col = .Rows(1).Find("端末矢崎品番", , , 1).Column
            cavCol = .Rows(1).Find("キャビティ", , , 1).Column
            For i = 2 To lastRow
                矢崎 = .Cells(i, 矢崎Col)
                cav = .Cells(i, cavCol)
                If 矢崎 <> "" Then
                    rs.filter = "(PartName = '" & 矢崎 & "') AND (Cav = '" & cav & "') AND (種類 = '" & "写真')"
                    If rs.EOF = True Then rs.filter = "(PartName = '" & 矢崎 & "') AND (Cav = '" & cav & "') AND (種類 = '" & "略図')"
                    For x = 1 To rs.fields.count
                        If addCol(x - 1) <> 0 Then
                            .Cells(i, addCol(x - 1)) = rs(x - 1)
                        End If
                    Next x
                End If
'                rs.Find "(PartName = '7283702640') AND (Cav = '1')", 0, adSearchForward
'                rs.Find "(PartName = '" & 矢崎 & "') AND (Cav = '" & Cav & "')", 0, adSearchForward
'                Do Until rs.EOF
'
'                Loop
            Next i
        End With
'        Dim J As Long: J = 0
'        Dim jj As Long: jj = 0
        'セルの値を削除
'        If a = LBound(mySQL) Then
'            'ワークシートの追加
'            If a = LBound(mySQL) Then
'                For Each ws In Worksheets
'                    If ws.Name = "ハメ図temp1" Then
'                        Application.DisplayAlerts = False
'                        ws.Delete
'                        Application.DisplayAlerts = True
'                    End If
'                Next ws
'                Set newSheet = Worksheets.Add
'                newSheet.Name = "ハメ図temp1"
'            End If
'        End If
        
'        J = 0
'        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
'                                       'PVSW_RLTFの書式設定を@にするとか
'        Do Until rs.EOF
'            ReDim Preserve 電線RAN(5, J)
'            For i = LBound(電線RAN, 1) To UBound(電線RAN, 1)
'                電線RAN(i, J) = rs(i)
'            Next i
'            J = J + 1
'            rs.MoveNext
'        Loop

'        With newSheet
'            .Cells.NumberFormat = "@"
'            For i = 0 To rs.Fields.count - 1
'                .Cells(1, i + 1) = Replace(Replace(rs(i).Name, "始点側", ""), "終点側", "")
'            Next i
'            lastRow = .UsedRange.Rows.count + 1
'            Debug.Print rs.RecordCount
'            .Cells(lastRow, 1).CopyFromRecordset rs
'        End With
        rs.Close
    Next a
    cn.Close
    
End Sub
Sub SQL_ローカル電線サブナンバー取得(ran, 製品品番)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    'ヘッダーの無いテキストファイルの時 12.0だとフィールド名がF1でとれない
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=NO;FMT=Delimited"
    cn.Open Left(myAddress(2, 1), InStrRev(myAddress(2, 1), "\") - 1)
    Set rs = New ADODB.Recordset

    Dim mysql(0) As String
    
    mysql(0) = " SELECT * " & _
          " FROM " & Mid(myAddress(2, 1), InStrRev(myAddress(2, 1), "\") + 1) & _
          " WHERE F1 = '" & 製品品番 & "' "
          
    For a = 0 To UBound(mysql)
        'SQLを開く
        'cn.Execute mySQL(0)
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        If a = 0 Then ReDim ran(rs.fields.count, 0): j = 0
        
        'Sheets("Sheet1").Cells.ClearContents
        Do Until rs.EOF
            ReDim Preserve ran(rs.fields.count, j)
            For i = 0 To rs.fields.count - 1
                'Sheets("Sheet1").Cells(J + 1, i + 1) = rs(i).Value
                ran(i, j) = rs(i).Value
            Next i
            j = j + 1
            'Range("a2").CopyFromRecordset rs
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Sub

Public Function SQL_ローカル端末サブナンバー取得(ran, 製品品番)
    
    If Dir(Left(myAddress(2, 1), InStrRev(myAddress(2, 1), "\") - 1) & "\TerminalSubNumber\" & Replace(製品品番, " ", "") & ".txt") = "" Then
        SQL_ローカル端末サブナンバー取得 = False
        Exit Function
    End If
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    'ヘッダーの無いテキストファイルの時 12.0だとフィールド名がF1でとれない
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "text;HDR=NO;FMT=Delimited"
    cn.Open Left(myAddress(2, 1), InStrRev(myAddress(2, 1), "\") - 1) & "\TerminalSubNumber\"
    Set rs = New ADODB.Recordset

    Dim mysql(0) As String
    
    mysql(0) = " SELECT * " & _
          " FROM " & Replace(製品品番, " ", "") & ".txt"
    
    For a = 0 To UBound(mysql)
        'SQLを開く
        'cn.Execute mySQL(0)
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        
        If a = 0 Then ReDim ran(rs.fields.count, 0): j = 0
        
        'Sheets("Sheet1").Cells.ClearContents
        Do Until rs.EOF
            ReDim Preserve ran(rs.fields.count, j)
            For i = 0 To rs.fields.count - 1
                'Sheets("Sheet1").Cells(J + 1, i + 1) = rs(i).Value
                ran(i, j) = rs(i).Value
            Next i
            j = j + 1
            'Range("a2").CopyFromRecordset rs
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Function
Sub SQL_マルマ変更(製品品番Ran, マルマ変更RAN, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).ActiveSheet
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "myTable"
    End With
    
    Dim mysql(0) As String
    
    mysql(0) = "SELECT "
    For i = 1 To 製品品番RANc
        mysql(0) = mysql(0) & "[" & 製品品番Ran(1, i - 1) & "],"
    Next i
    
    ReDim 電線一覧RAN(製品品番RANc + 9, 0)
    ReDim 端末一覧ran(0)
    mysql(0) = mysql(0) & "構成_,始点側回路符号, 終点側回路符号, 始点側端末識別子, 終点側端末識別子,始点側キャビティ,終点側キャビティ,線長_,線長後_ ,RLTFtoPVSW_,備考_" & _
          " FROM myTable " & _
          " WHERE " & "[RLTFtoPVSW_]='Found'" '& _
          " AND [線長後_] IS NOT NULL"

    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic, adLockReadOnly, 512
        Dim j As Long: j = 0
        Dim jj As Long: jj = 0
        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            '冶具typeの対象にあるか確認
            findFlg = False
            For i = 1 To 製品品番RANc
                If Not IsNull(rs(i - 1)) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            
            If findFlg = False Then
                GoTo line20
            End If
            
            '追加
            ReDim Preserve 電線一覧RAN(製品品番RANc + 9, j)
            
            For i = 1 To 製品品番RANc
                電線一覧RAN(i - 1, j + 0) = rs(i - 1)
            Next i
                '始点
                電線一覧RAN(製品品番RANc + 0, j + 0) = rs(製品品番RANc + 0) '構成
                電線一覧RAN(製品品番RANc + 1, j + 0) = rs(製品品番RANc + 1) '回符
                電線一覧RAN(製品品番RANc + 2, j + 0) = rs(製品品番RANc + 2)
                電線一覧RAN(製品品番RANc + 3, j + 0) = rs(製品品番RANc + 3) '端末
                電線一覧RAN(製品品番RANc + 4, j + 0) = rs(製品品番RANc + 4)
                電線一覧RAN(製品品番RANc + 5, j + 0) = rs(製品品番RANc + 5) 'cav
                電線一覧RAN(製品品番RANc + 6, j + 0) = rs(製品品番RANc + 6)
                電線一覧RAN(製品品番RANc + 7, j + 0) = rs(製品品番RANc + 7) '線長_
                電線一覧RAN(製品品番RANc + 8, j + 0) = rs(製品品番RANc + 8) '線長後_
                電線一覧RAN(製品品番RANc + 9, j + 0) = rs(製品品番RANc + 10) '備考_
                
            '始点端末無い時追加
            For i = LBound(端末一覧ran) To UBound(端末一覧ran)
                findFlg = False
                If 端末一覧ran(i) = rs(製品品番RANc + 3) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            If findFlg = False Then
                ReDim Preserve 端末一覧ran(jj)
                端末一覧ran(jj) = rs(製品品番RANc + 3)
                jj = jj + 1
            End If
            '終点端末無い時追加
            For i = LBound(端末一覧ran) To UBound(端末一覧ran)
                findFlg = False
                If 端末一覧ran(i) = rs(製品品番RANc + 4) Then
                    findFlg = True
                    Exit For
                End If
            Next i
            If findFlg = False Then
                ReDim Preserve 端末一覧ran(jj)
                端末一覧ran(jj) = rs(製品品番RANc + 4)
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

Sub SQL_互換端末cav(互換端末cavRAN, 互換端末RAN, 製品品番str, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
    cn.Open
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim 互換端末cavRAN(2, 0)
    Dim mysql(1) As String, 条件(4) As String
    '始点側の回路
    
    mysql(0) = " SELECT 始点側端末識別子,始点側キャビティ" & _
          " FROM 範囲 " & _
          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null and 始点側端末識別子 <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" '& _
          " GROUP BY 始点側端末識別子,始点側キャビティ"
    '終点側の回路
    mysql(1) = " SELECT 終点側端末識別子,終点側キャビティ" & _
          " FROM 範囲 " & _
          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null and 終点側端末識別子 <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" '& _
          " GROUP BY 終点側端末識別子,終点側キャビティ"
    Dim cnt As Long
    j = 0
    For a = 0 To 1
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        Do Until rs.EOF
            ReDim Preserve 互換端末cavRAN(2, j)
            
            For p = 0 To rs.fields.count - 1
                互換端末cavRAN(p, j) = rs(p)
            Next p
            For i = LBound(互換端末RAN, 2) To UBound(互換端末RAN, 2)
                If 互換端末RAN(0, i) = rs(0) Then
                    互換端末cavRAN(2, j) = 互換端末RAN(1, i)
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
Sub SQL_互換端末cav_1998(互換端末cavRAN, 互換端末RAN, 製品品番str, myBookName)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
    cn.Open
    Set rs = New ADODB.Recordset
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim 互換端末cavRAN(5, 0)
    Dim mysql(0) As String, 条件(4) As String
    '始点側の回路
    
    mysql(0) = " SELECT 始点側端末識別子,始点側キャビティ,終点側端末識別子,終点側キャビティ" & _
          " FROM 範囲 " & _
          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null" & _
          " AND " & "RLTFtoPVSW_='Found'" '& _
          " GROUP BY 始点側端末識別子,始点側キャビティ"
    '終点側の回路
'    mySQL(1) = " SELECT 終点側端末識別子,終点側キャビティ" & _
'          " FROM 範囲 " & _
'          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " <> Null and 終点側端末識別子 <> Null" & _
'          " AND " & "RLTFtoPVSW_='Found'" '& _
          " GROUP BY 終点側端末識別子,終点側キャビティ"
          
    Dim cnt As Long
    j = 0
    For a = 0 To 0
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        Do Until rs.EOF
            ReDim Preserve 互換端末cavRAN(5, j)
            
            For p = 0 To rs.fields.count - 1
                互換端末cavRAN(p, j) = rs(p)
            Next p
            
            Dim 始点flg As Boolean: 始点flg = False
            Dim 終点flg As Boolean: 終点flg = False
            For i = LBound(互換端末RAN, 2) To UBound(互換端末RAN, 2)
                '始点_端末が同じなら冶具座標をセット
                If 互換端末RAN(0, i) = rs(0) Then
                    互換端末cavRAN(4, j) = 互換端末RAN(1, i)
                    始点flg = True
                End If
                '終点_
                If 互換端末RAN(0, i) = rs(2) Then
                    互換端末cavRAN(5, j) = 互換端末RAN(1, i)
                    終点flg = True
                End If
                If 始点flg = True And 終点flg = True Then Exit For
            Next i
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close
    
End Sub

Sub SQL_配索回路取得(配索回路RAN, 製品品番str, サブstr)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
    cn.Open
    
    Dim コメント As String: コメント = "RLTFtoPVSW_" & " = " & "Found"
    
    With Sheets("PVSW_RLTF")
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    Set rs = New ADODB.Recordset
    
    Dim mysql As String
    mysql = " SELECT 色呼_,サ呼_,始点側端末識別子,始点側マ_,始点側ハメ,終点側端末識別子,終点側マ_,終点側ハメ" & _
          " FROM 範囲 " & _
          " WHERE " & Chr(34) & 製品品番str & Chr(34) & " = " & サブstr & " AND " & "RLTFtoPVSW_='Found'"   '& _
          " GROUP BY  始点側端末識別子,終点側端末識別子"

    'SQLを開く
    rs.Open mysql, cn, adOpenStatic
    '配列に格納
    ReDim 配索回路RAN(rs.fields.count - 1, rs.RecordCount - 1)
    Do Until rs.EOF
        For p = 0 To rs.fields.count - 1
            配索回路RAN(p, j) = rs(p)
        Next p
        j = j + 1
        rs.MoveNext
    Loop
    rs.Close
    cn.Close

End Sub
Sub SQL_配索サブ取得(配索サブRAN, 製品品番str)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With Sheets("端末一覧")
        Dim myRange As Range: Set myRange = .Cells.Find("端末矢崎品番", , , 1)
        Dim firstRow As Long: firstRow = myRange.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, myRange.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(myRange.Row, .Columns.count).End(xlToLeft).Column
        Set myRange = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    Set rs = New ADODB.Recordset
    
    Dim mysql As String
    mysql = " SELECT [" & 製品品番str & "] " & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
          " GROUP BY [" & 製品品番str & "]" & _
          " ORDER BY len([" & 製品品番str & "]),[" & 製品品番str & "]" ' & _
          " AND " & "RLTFtoPVSW_='Found'"   '& _
          " GROUP BY  始点側端末識別子,終点側端末識別子"

    'SQLを開く
    rs.Open mysql, cn, adOpenStatic
    '配列に格納
    ReDim 配索サブRAN(rs.fields.count - 1, rs.RecordCount - 1)
    Do Until rs.EOF
        For p = 0 To rs.fields.count - 1
            配索サブRAN(p, j) = rs(p)
        Next p
        j = j + 1
        rs.MoveNext
    Loop
    
    ReDim Preserve 配索サブRAN(0, UBound(配索サブRAN, 2) + 1)
    配索サブRAN(0, UBound(配索サブRAN, 2)) = "Base"
    rs.Close
    cn.Close

End Sub

Sub SQL_配索_端末経路取得(端末経路RAN, 製品品番str, 端末str)
      
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
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
        Dim 電線識別名 As Range: Set 電線識別名 = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = 電線識別名.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(電線識別名.Row, .Columns.count).End(xlToLeft).Column
        Set 電線識別名 = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    ReDim 端末経路RAN(6, 0)
    Dim mysql(1) As String, 条件(4) As String
    '始点側の回路
    mysql(0) = " SELECT 始点側端末識別子,終点側端末識別子, サ呼_,色呼_,終点側マ_,仕上寸法_,生区_" & _
          " FROM 範囲 " & _
          " WHERE [始点側端末識別子] = '" & 端末str & "'" & _
          " AND " & "RLTFtoPVSW_='Found'" & " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """""
    '終点側の回路
    mysql(1) = " SELECT 終点側端末識別子,始点側端末識別子, サ呼_,色呼_,始点側マ_,仕上寸法_,生区_" & _
          " FROM 範囲 " & _
          " WHERE [終点側端末識別子] = '" & 端末str & "'" & _
          " AND " & "RLTFtoPVSW_='Found'" & " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """""
    For a = 0 To 1
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        
        Do Until rs.EOF
            ReDim Preserve 端末経路RAN(rs.fields.count - 1, j)
            For p = 0 To rs.fields.count - 1
                端末経路RAN(p, j) = rs(p)
            Next p
            j = j + 1
            rs.MoveNext
        Loop
        
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_製品別端末一覧_使用電線確認(使用電線ran, 製品品番str)
    
    '参照設定= Microsoft ActiveX Data Objects 2.8 Library
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
        Dim myRange As Range: Set myRange = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = myRange.Row
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, myRange.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(myRange.Row, .Columns.count).End(xlToLeft).Column
        Set myRange = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
    End With
    
    Set rs = New ADODB.Recordset
    Dim mysql(1) As String
    'Dim 使用電線ran()
    ReDim 使用電線ran(3, 0)
    j = 0
    mysql(0) = " SELECT [" & 製品品番str & "],[始点側端末識別子] ,[始点側端末矢崎品番],[始点側キャビティ]" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
          " AND " & "RLTFtoPVSW_='Found'"   '& _
          " GROUP BY  始点側端末識別子,終点側端末識別子"
    mysql(1) = " SELECT [" & 製品品番str & "] ,[終点側端末識別子],[終点側端末矢崎品番],[終点側キャビティ]" & _
          " FROM 範囲 " & _
          " WHERE [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
          " AND " & "RLTFtoPVSW_='Found'"
    For a = LBound(mysql) To UBound(mysql)
          '& _
              " GROUP BY  終点側端末識別子,終点側端末識別子"
        'SQLを開く
        rs.Open mysql(a), cn, adOpenStatic
        '使用しているCAVを格納
        Do Until rs.EOF
            ReDim Preserve 使用電線ran(rs.fields.count - 1, j)
            For p = 0 To rs.fields.count - 1
                使用電線ran(p, j) = rs(p)
            Next p
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
    Next a
    cn.Close

End Sub

Sub SQL_配策図用_製品品番_構成_SUB(ran, 製品品番str, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(1, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
    For k = 0 To 0
        mysql(0) = " SELECT [" & 製品品番str & "],構成_,'" & 製品品番str & "'" & _
              " FROM 範囲 " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
              " ORDER BY [" & 製品品番str & "] ASC"
'        mySQL(1) = " SELECT [" & 製品品番str & "],終点側端末矢崎品番,終点側端末識別子,TI2,'" & 製品品番str & "'" & _
'              " FROM 範囲 " & _
'              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
'              " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
'              " ORDER BY [" & 製品品番str & "] ASC"
    
        'SQLを開く
        rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            flg = False
            '追加
            If flg = False Then
                If rs(1) & rs(0) <> "" Then
                    j = j + 1
                    ReDim Preserve ran(1, j)
                    ran(0, j) = Replace(rs(2), " ", "") & "_" & rs(1) '製品品番_構成
                    ran(1, j) = rs(0) 'Sub
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next k
    cn.Close

End Sub

Public Function SQL_電線情報RANset(ran, 製品品番str, myBook, 端末)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String: xl_file = ThisWorkbook.FullName
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    With myBook.Sheets("PVSW_RLTF両端")
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With

    SQL_電線情報RANset = 0
    ReDim ran(9, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
    For k = 0 To 0
        mysql(0) = " SELECT [" & 製品品番str & "],端末矢崎品番,サ呼_,切断長_,端子_,端末識別子,構成_,マ_,相手_,'" & 製品品番str & "'" & _
              " FROM 範囲 " & _
              " WHERE " & "[RLTFtoPVSW_]='Found' AND [端末識別子]='" & 端末 & "'" & _
              " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
              " ORDER BY 切断長_ DESC, 相手_ ASC"  '←このソートきいてない
'        mySQL(1) = " SELECT [" & 製品品番str & "],終点側端末矢崎品番,終点側端末識別子,TI2,'" & 製品品番str & "'" & _
'              " FROM 範囲 " & _
'              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
'              " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
'              " ORDER BY [" & 製品品番str & "] ASC"
        'SQLを開く
        rs.CursorLocation = adUseClient
        rs.Open mysql(k), cn, adOpenKeyset, adLockOptimistic, 512
        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            flg = False
            If flg = False Then
                If rs(1) & rs(0) <> "" Then
                    If Left(rs(4), 4) = "7409" Or Left(rs(4), 4) = "7009" Then
                        j = j + 1
                        ReDim Preserve ran(9, j)
                        For i = 0 To rs.fields.count - 1
                            If Not IsNull(rs(i)) Then
                                ran(i, j) = Replace(rs(i), " ", "")
                            End If
                        Next i
                        SQL_電線情報RANset = j
                    End If
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next k
    cn.Close
    
    If j > 0 Then
        '切断長_順で並び替えする
        Dim myAry1()
        myAry1 = WorksheetFunction.transpose(ran) 'SQLでセットした配列を入れ替える
        '2次元バブルソート
        Call BubbleSort2(myAry1, 4) '昇順
        ran = WorksheetFunction.transpose(myAry1)
    End If
End Function

Sub SQL_配策図用_回路(ran, 製品品番str, myBook)
    
    '参照設定= Microsoft ActiveX Data Objects 6.1 Library
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
        Dim key As Range: Set key = .Cells.Find("電線識別名", , , 1)
        Dim firstRow As Long: firstRow = key.Row
        Dim lastRow As Long
        lastRow = .UsedRange.Rows.count '.Cells(.Rows.count, 電線識別名.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Set key = Nothing
        .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol)).Name = "範囲"
        Set myTitle = .Range(.Cells(firstRow, 1), .Cells(firstRow, lastCol))
    End With
    
    ReDim ran(11, 0): j = 0
    Dim mysql() As String: ReDim mysql(0)
    For k = 0 To 0
        mysql(0) = " SELECT [" & 製品品番str & "],構成_,色呼_,始点側端末識別子,終点側端末識別子,'" & 製品品番str & "'" & ",始点側ハメ,始点側キャビティ,終点側ハメ,終点側キャビティ" & _
              " FROM 範囲 " & _
              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
              " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
              " ORDER BY [" & 製品品番str & "] ASC"
'        mySQL(1) = " SELECT [" & 製品品番str & "],終点側端末矢崎品番,終点側端末識別子,TI2,'" & 製品品番str & "'" & _
'              " FROM 範囲 " & _
'              " WHERE " & "[RLTFtoPVSW_]='Found'" & _
'              " AND [" & 製品品番str & "] IS NOT NULL AND [" & 製品品番str & "] <> """"" & _
'              " ORDER BY [" & 製品品番str & "] ASC"
    
        'SQLを開く
        rs.Open mysql(k), cn, adOpenStatic, adLockReadOnly, 512
        If rs(0).Type <> 202 Then Stop 'rsの製品品番strのデータタイプが202じゃないから文字列が抜ける
                                       'PVSW_RLTFの書式設定を@にするとか
        Do Until rs.EOF
            flg = False
            '追加
            If flg = False Then
                If rs(1) & rs(0) <> "" Then
                    j = j + 1
                    ReDim Preserve ran(11, j)
                    ran(0, j) = Replace(rs(5), " ", "") '製品品番
                    ran(1, j) = rs(0) 'Sub
                    ran(2, j) = rs(1)
                    ran(3, j) = rs(2)
                    ran(4, j) = rs(3)
                    ran(5, j) = rs(4)
                    ran(6, j) = rs(5)
                    ran(7, j) = rs(6)
                    ran(8, j) = rs(7)
                    ran(9, j) = rs(8)
                    ran(10, j) = rs(9)
                    ran(11, j) = 色変換(rs(2), clocode1, clocode2, clofont) '色呼long
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next k
    cn.Close

End Sub

Sub SQLもどき_後ハメ作業者(ran, 製品品番str)

    'シート名が大きいシートの検索
    Dim wsTemp As Worksheet, wsNumber As Long
    For Each wsTemp In wb(3).Worksheets
        If IsNumeric(wsTemp.Name) Then
            If CLng(wsTemp.Name) > wsNumber Then
                wsNumber = wsTemp.Name
            End If
        End If
    Next wsTemp

    If wsNumber = 0 Then
        MsgBox "シート名に数字が見つかりません。中断します"
        Call 最適化もどす
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
        .Range(.Cells(firstRow, 2), .Cells(lastRow, lastCol)).Name = "範囲"
    End With

    With wb(3).Sheets(CStr(wsNumber))
        Dim 製品品番check As Variant
        Set 製品品番check = .Rows(firstRow).Find(製品品番str, , , 1)
        If 製品品番check Is Nothing Then MsgBox 製品品番str & "が後嵌め作業者一覧表にありません。中断します。": End
        Dim Col0 As Long: Col0 = .Rows(firstRow).Find("key_", , , 1).Column
        Dim Col1 As Long: Col1 = 製品品番check.Column
        ReDim ran(4, 0) '3,4は端末,cav入れる用
        ran(0, 0) = "構成"
        ran(1, 0) = "後ハメ作業者"
        ran(2, 0) = "作業順"
        ran(3, 0) = "端末識別子"
        ran(4, 0) = "cav"
        Dim 後ハメ作業者str  As String, 構成 As String, ハメ作業順 As String
        C = 0
        For y = koseiRow + 1 To lastRow
            構成 = .Cells(y, Col0)
            If 構成 <> "" Then
                後ハメ作業者str = .Cells(y, Col1)
                ハメ作業順 = .Cells(y, Col1).Offset(0, -1)
                ReDim Preserve ran(UBound(ran), UBound(ran, 2) + 1)
                ran(0, UBound(ran, 2)) = 構成
                ran(1, UBound(ran, 2)) = 後ハメ作業者str
                ran(2, UBound(ran, 2)) = ハメ作業順
            End If
        Next y
        Set myKey = Nothing
    End With

    後ハメ作業者シート名 = wsNumber & "版"
End Sub
