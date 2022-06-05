VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_08 
   Caption         =   "サブ立案"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
   OleObjectBlob   =   "UI_08.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
















































































Private Sub CB0_Change()
    Dim 項目(1) As String
    Dim 項目2(1) As String
    'CB0.Text
    With ActiveWorkbook.Sheets("製品品番")
        Set myKey = .Cells.Find("型式", , , 1)
        Set myKey = .Rows(myKey.Row).Find(CB0.Text, , , 1)
        Set mykey2 = .Rows(myKey.Row).Find("結き", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For y = myKey.Row + 1 To lastRow
            If InStr(項目(0), "," & .Cells(y, myKey.Column)) & "," = 0 Then
                項目(0) = 項目(0) & "," & .Cells(y, myKey.Column) & ","
                項目2(0) = 項目2(0) & "," & .Cells(y, mykey2.Column) & ","
            End If
        Next y
        If Len(項目(0)) <= 2 Then
            項目(0) = ""
            項目s = Empty
        Else
            項目(0) = Mid(項目(0), 2)
            項目(0) = Left(項目(0), Len(項目(0)) - 1)
            項目s = Split(項目(0), ",,")
            項目2(0) = Mid(項目2(0), 2)
            項目2(0) = Left(項目2(0), Len(項目2(0)) - 1)
            項目2s = Split(項目2(0), ",,")
        End If
    End With
    
    With CB1
        .RowSource = ""
        .Clear
        If Not IsEmpty(項目s) Then
            For i = LBound(項目s) To UBound(項目s)
                .AddItem
                .List(i, 0) = 項目s(i)
                .List(i, 1) = 項目2s(i)
            Next i
            .ListIndex = 0
        End If
    End With
End Sub

Private Sub CB1_Change()
    Call 製品品番RAN_set2(製品品番Ran, CB0.Value, CB1.Value, "")
    If 製品品番RANc <> 1 Then
        myLabel.Caption = "製品品番点数が異常です。"
        myLabel.ForeColor = RGB(255, 0, 0)
        Exit Sub
    Else
        myLabel.Caption = ""
    End If
End Sub

Private Sub CommandButton4_Click()
    PlaySound "もどる"
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    mytime = time
    PlaySound "じっこう"
    Call 製品品番RAN_set2(製品品番Ran, CB0.Value, CB1.Value, "")

    Unload Me
    Call checkSheet("PVSW_RLTF;端末一覧", wb(0), True, True)
    
    Call PVSWcsv両端のシート作成_Ver2001
    Call PVSWcsvにサブナンバーを渡してサブ図データ作成_2017
    
    '使用するフィールド名のセット
    Dim fieldname As String: fieldname = "RLTFtoPVSW_,始点側端末識別子,終点側端末識別子,始点側端末矢崎品番,終点側端末矢崎品番,仕上寸法_,接続G_,構成_,生区_"
    ff = Split(fieldname, ",")
    Dim f As Variant: ReDim f(UBound(ff))
    For x = LBound(ff) To UBound(ff)
        f(x) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(x), , , 1).Column
    Next x
    a = UBound(ff) + 1
    '電線数をセットする配列
    Dim 端末電線数RAN As Variant
    ReDim 端末電線数RAN(a, 0)
    'フィールド名を配列に入れる
    For x = LBound(ff) To UBound(ff)
        端末電線数RAN(x, 0) = ff(x)
    Next x
    端末電線数RAN(UBound(端末電線数RAN), 0) = "親端末No"
    
    '対象のグループ毎に処理
    Dim メイン品番i As Integer
    メイン品番i = 製品品番RAN_read(製品品番Ran, "メイン品番")
    For i = LBound(製品品番Ran, 2) + 1 To UBound(製品品番Ran, 2)
        製品品番str = 製品品番Ran(メイン品番i, i)
        With wb(0).Sheets("PVSW_RLTF")
            '製品品番のフィールドをキーとしてセット
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(製品品番str, , , 1)
            Dim lastRow As Long
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For y = myKey.Row + 1 To lastRow
                If .Cells(y, myKey.Column) <> "" Then
                    If .Cells(y, f(0)) = "Found" Then
                        ReDim Preserve 端末電線数RAN(a, UBound(端末電線数RAN, 2) + 1)
                        For x = LBound(f) + 1 To UBound(f)
                            端末電線数RAN(x, UBound(端末電線数RAN, 2)) = .Cells(y, f(x))
                        Next x
                        '端末電線数RAN(0, UBound(端末電線数RAN, 2)) = 1
                    End If
                End If
            Next y
        End With
       Call ReplaceLR(端末電線数RAN)
    
        Call SumRan(端末電線数RAN) '両端の行き先が同じ、接続Gが同じ場合まとめる
        '回路の多さから親を決める
        Dim 端末評価RAN()
        端末評価RAN = evaluationRan(端末電線数RAN) '優先が9の時サブナンバー999
        端末評価RAN = changeRowCol(端末評価RAN)
        Call BubbleSort3(端末評価RAN, 3, 2)
        端末電線数RAN = changeRowCol(端末電線数RAN)
    Next i
    
    '端末サブナンバー999の電線サブナンバーを999にする
    For i = LBound(端末評価RAN) To UBound(端末評価RAN)
        端末str = 端末評価RAN(i, 0)
        サブstr = 端末評価RAN(i, 5)
        If サブstr <> "" Then
            For ii = LBound(端末電線数RAN) To UBound(端末電線数RAN)
                For x = 1 To 2
                    If 端末str = 端末電線数RAN(ii, x) Then
                        端末電線数RAN(ii, UBound(端末電線数RAN, 2)) = サブstr
                        Exit For
                    End If
                Next x
            Next ii
        End If
    Next i
    
    'Call export_ArrayToSheet(端末電線数RAN, "端末電線数RAN", False)
    'todo
    'サブナンバーを配布
    For ii = LBound(端末電線数RAN) To UBound(端末電線数RAN)
        '接続Gによる判断
        Select Case Left(端末電線数RAN(ii, 6), 1)
            Case "T"
                '何もしない
            Case "E", "J", "B"
                端末電線数RAN(ii, UBound(端末電線数RAN, 2)) = "999"
            Case "W"
                端末電線数RAN(ii, UBound(端末電線数RAN, 2)) = "999"
        End Select
        '生区_による判断
        Select Case Left(端末電線数RAN(ii, 8), 1)
            Case "E"
                端末電線数RAN(ii, UBound(端末電線数RAN, 2)) = "999"
        End Select
    Next ii
    'Call export_ArrayToSheet(端末電線数RAN, "端末電線数RAN", False)
    
    '評価の高い親を基準にサブナンバーを配布していく
    For i = LBound(端末評価RAN) + 1 To UBound(端末評価RAN)
        端末str = 端末評価RAN(i, 0)
        相手端末数str = 端末評価RAN(i, 6)
        'If 端末str = "250" Then Stop
        If 端末評価RAN(i, 5) <> "" Then GoTo line20
        端末評価RAN(i, 5) = 端末str
        For j = LBound(端末電線数RAN) + 1 To UBound(端末電線数RAN)
'            If j = 227 Then Stop
'            If i = 3 And j = 207 Then Stop
            If 端末電線数RAN(j, UBound(端末電線数RAN, 2)) = "" Then 'まだサブナンバーが決まって無ければ
                For x = 1 To 2
                    If 端末str = 端末電線数RAN(j, x) Then
                        端末評価lng = 端末電線数RAN(j, 0)
                        If x = 1 Then 相手端末str = 端末電線数RAN(j, 2)
                        If x = 2 Then 相手端末str = 端末電線数RAN(j, 1)
                        'もし相手端末数が1の場合に相手端末のサブナンバーに変更
                        If 相手端末数str = "1" Then
                            相手端末サブstr = search端末評価RAN(端末評価RAN, 相手端末str, 5)
                            If 相手端末サブstr <> "" Then
                                端末評価RAN(i, 5) = 相手端末サブstr
                                GoTo line20
                            End If
                        End If
                        相手端末優先 = search端末評価RAN(端末評価RAN, 相手端末str, 3)
                        If 相手端末優先 = "1" Then GoTo line15
                        相手端末評価lng = search相手端末評価(端末電線数RAN, 相手端末str)
                        If 端末評価lng >= 相手端末評価lng Then
                            If 端末電線数RAN(j, UBound(端末電線数RAN, 2)) = "" Then
                                端末電線数RAN(j, UBound(端末電線数RAN, 2)) = 端末str
                            End If
                            For ii = LBound(端末評価RAN) + 1 To UBound(端末評価RAN)
                                If 端末評価RAN(ii, 0) = 相手端末str Then
                                    If 端末評価RAN(ii, 5) = "" Then
                                        端末評価RAN(ii, 5) = 端末str
                                    End If
                                    Exit For
                                End If
                            Next ii
                        End If
                    End If
line15:
                Next x
            End If
        Next j
line20:
    Next i
    
    'Call export_ArrayToSheet(端末評価RAN, "端末評価RAN", False)
    
    'Call export_ArrayToSheet(端末電線数RAN, "端末電線数RAN", False)
    '繋がらなかった電線を評価の高い端末に挿すようにする
    For i = LBound(端末電線数RAN) + 1 To UBound(端末電線数RAN)
        If 端末電線数RAN(i, 0) <> "" Then
            If 端末電線数RAN(i, UBound(端末電線数RAN, 2)) = "" Then
                端末1str = 端末電線数RAN(i, 1)
                端末2str = 端末電線数RAN(i, 2)
                端末評価1str = search端末評価RAN(端末評価RAN, 端末1str, 2)
                端末評価2str = search端末評価RAN(端末評価RAN, 端末2str, 2)
                If 端末評価1str > 端末評価2str Then
                    所属サブstr = search端末評価RAN(端末評価RAN, 端末1str, 5)
                Else
                    所属サブstr = search端末評価RAN(端末評価RAN, 端末2str, 5)
                End If
                端末電線数RAN(i, UBound(端末電線数RAN, 2)) = 所属サブstr
            End If
        End If
    Next i
    
    'Call export_ArrayToSheet(端末電線数RAN, "端末電線数RAN", False)
    
    '端末888の配布
    Dim 端末str1 As String, 端末str2 As String, 接続Gstr As String
    For i = LBound(端末電線数RAN) + 1 To UBound(端末電線数RAN)
        端末str1 = 端末電線数RAN(i, 1)
        端末str2 = 端末電線数RAN(i, 2)
        If 端末str1 & 端末str2 = "" Then 端末電線数RAN(i, UBound(端末電線数RAN, 2)) = "999"
    Next i
    
'    Call export_ArrayToSheet(端末電線数RAN, "端末電線数RAN", False)
    
    'RLFTtoPVSW_が空欄の場合除外する
    For i = LBound(端末電線数RAN) To UBound(端末電線数RAN)
        If i > UBound(端末電線数RAN) Then Exit For
        If 端末電線数RAN(i, 0) = "" Then
            端末電線数RAN = removeArrayIndex(端末電線数RAN, i)
        End If
    Next i
    
   'Call export_ArrayToSheet(端末電線数RAN, "端末電線数RAN", False)
    
    '端末のサブナンバーが電線のサブナンバーに無い場合、cにする
    Dim foundFlg As Boolean
    For i = LBound(端末評価RAN) + 1 To UBound(端末評価RAN)
        foundFlg = False
        サブstr = 端末評価RAN(i, 5)
        For ii = LBound(端末電線数RAN) + 1 To UBound(端末電線数RAN)
            If サブstr = 端末電線数RAN(ii, UBound(端末電線数RAN, 2)) Then
                foundFlg = True
                Exit For
            End If
        Next ii
        If foundFlg = False Then
            端末評価RAN(i, 5) = "c"
        End If
    Next i
    
    '端末評価RANをテキスト出力
    Dim myTextPath As String
    myTextPath = wb(0).path & dirString_09
    makeDir myTextPath
    myTextPath = myTextPath & Replace(製品品番str, " ", "") & "_term.txt"
    export_Array_ShiftJis 端末評価RAN, myTextPath, ","
    
    '端末電線数RANをテキスト出力
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(製品品番str, " ", "") & "_wiresum.txt"
    export_Array_ShiftJis 端末電線数RAN, myTextPath, ","
    
    'Call export_ArrayToSheet(端末電線数RAN, "端末電線数RAN", False)
    
    '電線毎にサブナンバー(作業順)とステップナンバーを決める
    Dim myRan As Variant
    myRan = setWorkRanV2(製品品番str)
    
'    Call export_ArrayToSheet(端末評価RAN, "端末評価RAN", False)
    
    '電線毎のサブナンバーを端末毎のサブナンバーに渡す
    Dim subNumber As String, 親端末str As String
    For y = LBound(端末評価RAN) To UBound(端末評価RAN)
        親端末str = 端末評価RAN(y, 5)
        If 親端末str = "c" Then
            subNumber = "c"
        Else
            subNumber = searchRan_ver2(myRan, 親端末str, "親端末No", "subNumber")
        End If
        端末評価RAN(y, UBound(端末評価RAN, 2)) = subNumber
    Next y
    端末評価RAN = WorksheetFunction.transpose(端末評価RAN)
    
    'Call export_ArrayToSheet(myRan, "myRan", True)
    
    'PVSW_RLTFにサブナンバーを配布
    For i = LBound(製品品番Ran, 2) + 1 To UBound(製品品番Ran, 2)
        製品品番str = 製品品番Ran(メイン品番i, i)
        With wb(0).Sheets("PVSW_RLTF")
            '製品品番のフィールドをキーとしてセット
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(製品品番str, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For y = myKey.Row + 1 To lastRow
                If .Cells(y, myKey.Column) <> "" Then
                    If .Cells(y, f(0)) = "Found" Then
                        構成str = Left(.Cells(y, f(7)), 4)
                        subNumber = searchRan_ver2(myRan, 構成str, "構成_", "subNumber")
                        .Cells(y, myKey.Column) = subNumber
                        .Cells(y, myKey.Column).Interior.color = theme_color1
                    End If
                End If
            Next y
        End With
    Next i
    
    '端末一覧にサブナンバーを配布
    For i = LBound(製品品番Ran, 2) + 1 To UBound(製品品番Ran, 2)
        製品品番str = 製品品番Ran(メイン品番i, i)
        Set ws(3) = wb(0).Sheets("端末一覧")
        With ws(3)
            Dim myCol(1) As Integer
            myCol(0) = .Cells.Find("端末矢崎品番", , , 1).Column
            myCol(1) = .Cells.Find("端末№", , , 1).Column
            '製品品番のフィールドをキーとしてセット
            Set myKey = ws(3).Cells.Find(製品品番str, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For y = myKey.Row + 1 To lastRow
                If .Cells(y, myKey.Column) <> "" Then
                    端末矢崎品番str = .Cells(y, myCol(0))
                    端末str = .Cells(y, myCol(1))
                    subNumber = searchRan_ver2(端末評価RAN, 端末str & "," & 端末矢崎品番str, "端末No,端末矢崎品番", "subNumber")
                    If subNumber = "" Then Stop
                    .Cells(y, myKey.Column) = subNumber
                    .Cells(y, myKey.Column).Interior.color = theme_color1
                End If
            Next y
        End With
    Next i
    
    '配列をテキストファイル出力する
    '端末評価RANをテキスト出力
    端末評価RAN = WorksheetFunction.transpose(端末評価RAN)
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(製品品番str, " ", "") & "_term.txt"
    export_Array_ShiftJis 端末評価RAN, myTextPath, ","
    
    '端末電線数RANをテキスト出力
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(製品品番str, " ", "") & "_wireSum.txt"
    export_Array_ShiftJis 端末電線数RAN, myTextPath, ","
    
    'myRANをテキスト出力
    myRan = WorksheetFunction.transpose(myRan)
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(製品品番str, " ", "") & "_wire.txt"
    export_Array_ShiftJis myRan, myTextPath, ","
    
    Call 最適化もどす
    PlaySound "かんせい"
'
'    addRow = 1
'    For y = LBound(端末電線数RAN) To UBound(端末電線数RAN)
'            For x = LBound(端末電線数RAN, 2) To UBound(端末電線数RAN, 2)
'                With Sheets("temp")
'                    .Cells(addRow, x + 1) = 端末電線数RAN(y, x)
'                End With
'            Next x
'            addRow = addRow + 1
'    Next y
'
'    addRow = 1
'    For y = LBound(端末評価RAN) To UBound(端末評価RAN)
'        For x = LBound(端末評価RAN, 2) To UBound(端末評価RAN, 2)
'            With Sheets("temp2")
'                .Cells(addRow, x + 1) = 端末評価RAN(y, x)
'            End With
'        Next x
'        addRow = addRow + 1
'    Next y

    Dim myMsg As String: myMsg = "処理しました" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "生産準備+サブ自動立案")
End Sub

Private Sub UserForm_Initialize()
    Dim 項目(1) As String
    With ActiveWorkbook.Sheets("製品品番")
        Set myKey = .Cells.Find("型式", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For x = myKey.Column To lastCol
            項目(0) = 項目(0) & "," & .Cells(myKey.Row, x)
        Next x
        項目(0) = Mid(項目(0), 2)
    End With
    項目s = Split(項目(0), ",")
    With CB0
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
            If 項目s(i) = "メイン品番" Then myindex = i
        Next i
        .ListIndex = myindex
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "とじる"
End Sub
