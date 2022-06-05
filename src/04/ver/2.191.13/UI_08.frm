VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_08 
   Caption         =   "サブ立案"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
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
        For Y = myKey.Row + 1 To lastRow
            If InStr(項目(0), "," & .Cells(Y, myKey.Column)) & "," = 0 Then
                項目(0) = 項目(0) & "," & .Cells(Y, myKey.Column) & ","
                項目2(0) = 項目2(0) & "," & .Cells(Y, mykey2.Column) & ","
            End If
        Next Y
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
    Call 製品品番RAN_set2(製品品番RAN, CB0.Value, CB1.Value, "")
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
    mytime = Time
    PlaySound "じっこう"
    Call 製品品番RAN_set2(製品品番RAN, CB0.Value, CB1.Value, "")

    Unload Me
    Set wb(0) = ActiveWorkbook
    
    '使用するフィールド名のセット
    Dim fieldName As String: fieldName = "RLTFtoPVSW_,始点側端末識別子,終点側端末識別子,始点側端末矢崎品番,終点側端末矢崎品番,仕上寸法_"
    ff = Split(fieldName, ",")
    Dim f As Variant: ReDim f(UBound(ff))
    For X = LBound(ff) To UBound(ff)
        f(X) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(X), , , 1).Column
    Next X
    a = UBound(ff) + 1
    '電線数をセットする配列
    Dim 端末電線数RAN As Variant
    ReDim 端末電線数RAN(a, 0)
    'フィールド名を配列に入れる
    For X = LBound(ff) To UBound(ff)
        端末電線数RAN(X, 0) = ff(X)
    Next X
    '対象のグループ毎に処理
    Dim メイン品番i As Integer
    メイン品番i = 製品品番RAN_read(製品品番RAN, "メイン品番")
    For i = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
        製品品番str = 製品品番RAN(メイン品番i, i)
        With wb(0).Sheets("PVSW_RLTF")
            '製品品番のフィールドをキーとしてセット
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(製品品番str, , , 1)
            Dim lastRow As Long
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For Y = myKey.Row + 1 To lastRow
                If .Cells(Y, myKey.Column) <> "" Then
                    If .Cells(Y, f(0)) = "Found" Then
                        ReDim Preserve 端末電線数RAN(a, UBound(端末電線数RAN, 2) + 1)
                        For X = LBound(f) + 1 To UBound(f)
                            端末電線数RAN(X, UBound(端末電線数RAN, 2)) = .Cells(Y, f(X))
                        Next X
                        '端末電線数RAN(0, UBound(端末電線数RAN, 2)) = 1
                    End If
                End If
            Next Y
        End With
       Call ReplaceLR(端末電線数RAN)
       Call SumRan(端末電線数RAN)
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
                For X = 1 To 2
                    If 端末str = 端末電線数RAN(ii, X) Then
                        端末電線数RAN(ii, 6) = サブstr
                        Exit For
                    End If
                Next X
            Next ii
        End If
    Next i
    
    '評価の高い親を基準にサブナンバーを配布していく
    For i = LBound(端末評価RAN) + 1 To UBound(端末評価RAN)
        端末str = 端末評価RAN(i, 0)
        相手端末数str = 端末評価RAN(i, 6)
        'If 端末str = "250" Then Stop
        If 端末評価RAN(i, 5) <> "" Then GoTo line20
        端末評価RAN(i, 5) = 端末str
        For j = LBound(端末電線数RAN) + 1 To UBound(端末電線数RAN)
            If 端末電線数RAN(j, 6) = "" Then 'まだサブナンバーが決まって無ければ
                For X = 1 To 2
                    If 端末str = 端末電線数RAN(j, X) Then
                        端末評価lng = 端末電線数RAN(j, 0)
                        If X = 1 Then 相手端末str = 端末電線数RAN(j, 2)
                        If X = 2 Then 相手端末str = 端末電線数RAN(j, 1)
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
                            端末電線数RAN(j, 6) = 端末str
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
                Next X
            End If
        Next j
line20:
    Next i
    
    '繋がらなかった電線を評価の高い端末に挿すようにする
    For i = LBound(端末電線数RAN) + 1 To UBound(端末電線数RAN)
        If 端末電線数RAN(i, 0) <> "" Then
            If 端末電線数RAN(i, 6) = "" Then
                端末1str = 端末電線数RAN(i, 1)
                端末2str = 端末電線数RAN(i, 2)
                端末評価1str = search端末評価RAN(端末評価RAN, 端末1str, 2)
                端末評価2str = search端末評価RAN(端末評価RAN, 端末2str, 2)
                If 端末評価1str > 端末評価2str Then
                    所属サブstr = search端末評価RAN(端末評価RAN, 端末1str, 5)
                Else
                    所属サブstr = search端末評価RAN(端末評価RAN, 端末2str, 5)
                End If
                端末電線数RAN(i, 6) = 所属サブstr
            End If
        End If
    Next i
    
    addRow = 1
    For Y = LBound(端末電線数RAN) To UBound(端末電線数RAN)
        If 端末電線数RAN(Y, 0) <> "" Then
            For X = LBound(端末電線数RAN, 2) To UBound(端末電線数RAN, 2)
                With Sheets("temp")
                    .Cells(addRow, X + 1) = 端末電線数RAN(Y, X)
                End With
            Next X
            addRow = addRow + 1
        End If
    Next Y
    
    addRow = 1
    For Y = LBound(端末評価RAN) To UBound(端末評価RAN)
            For X = LBound(端末評価RAN, 2) To UBound(端末評価RAN, 2)
                With Sheets("temp2")
                    .Cells(addRow, X + 1) = 端末評価RAN(Y, X)
                End With
            Next X
            addRow = addRow + 1
    Next Y
    Stop
    
    'PVSW_RLTFにサブナンバーを配布
    For i = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
        製品品番str = 製品品番RAN(メイン品番i, i)
        With wb(0).Sheets("PVSW_RLTF")
            '製品品番のフィールドをキーとしてセット
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(製品品番str, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For Y = myKey.Row + 1 To lastRow
                If .Cells(Y, myKey.Column) <> "" Then
                    If .Cells(Y, f(0)) = "Found" Then
                        端末str1 = .Cells(Y, f(1))
                        端末str2 = .Cells(Y, f(2))
                        '端末str1を小さい数字に揃える
                        swapflg = False
                        If 端末str1 = "" Then swapflg = True
                        If IsNumeric(端末str1) = True And IsNumeric(端末str2) = True Then
                            If Val(端末str1) > Val(端末str2) Then
                                swapflg = True
                            End If
                        End If
                        If swapflg = True Then
                            vSwap = 端末str2
                            端末str2 = 端末str1
                            端末str1 = vSwap
                        End If
                        If 端末str1 & 端末str2 <> "" Then
                            サブstr = search端末電線数RAN(端末電線数RAN, 端末str1, 端末str2, 6)
                            If サブstr = "" Then Stop '担当者に連絡
                            .Cells(Y, myKey.Column) = サブstr
                            .Cells(Y, myKey.Column).Interior.color = RGB(129, 216, 208)
                        End If
                    End If
                End If
            Next Y
        End With
    Next i
    
    '端末一覧にサブナンバーを配布
    For i = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
        製品品番str = 製品品番RAN(メイン品番i, i)
        Set ws(3) = wb(0).Sheets("端末一覧")
        With ws(3)
            Dim myCol(1) As Integer
            myCol(0) = .Cells.Find("端末矢崎品番", , , 1).Column
            myCol(1) = .Cells.Find("端末№", , , 1).Column
            '製品品番のフィールドをキーとしてセット
            Set myKey = ws(3).Cells.Find(製品品番str, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For Y = myKey.Row + 1 To lastRow
                If .Cells(Y, myKey.Column) <> "" Then
                    端末矢崎品番str = .Cells(Y, myCol(0))
                    端末str = .Cells(Y, myCol(1))
                    サブstr = search端末評価RAN_2pos(端末評価RAN, 端末str, 端末矢崎品番str, 5)
                    If サブstr = "" Then Stop
                    .Cells(Y, myKey.Column) = サブstr
                    .Cells(Y, myKey.Column).Interior.color = RGB(129, 216, 208)
                End If
            Next Y
        End With
    Next i
    
    Stop
    
    Call 最適化もどす
    PlaySound "かんせい"
    
    Dim myMsg As String: myMsg = "作成しました" & vbCrLf & DateDiff("s", mytime, Time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "生産準備+配索誘導")
End Sub

Private Sub UserForm_Initialize()
    Dim 項目(1) As String
    With ActiveWorkbook.Sheets("製品品番")
        Set myKey = .Cells.Find("型式", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For X = myKey.Column To lastCol
            項目(0) = 項目(0) & "," & .Cells(myKey.Row, X)
        Next X
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
