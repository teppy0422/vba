VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_01 
   Caption         =   "ハメ図作成"
   ClientHeight    =   8430
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9580
   OleObjectBlob   =   "UI_01.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





















































Private Sub B0_Click()
    PlaySound "せんたく"
    CB0.ListIndex = 4
    CB1.ListIndex = 1
    CB2.ListIndex = 1
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 1
    cbx0.Value = True
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
    'PIC00.Picture = LoadPicture(myaddress(0,1) & "\ハメ図sample_" & "4511000000" & ".jpg")
End Sub

Private Sub B1_Click()
    PlaySound "せんたく"
    CB0.ListIndex = 1
    CB1.ListIndex = 4
    CB2.ListIndex = 1
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 1
    cbx0.Value = True
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
End Sub

Private Sub B2_Click()
    PlaySound "せんたく"
    CB0.ListIndex = 2
    CB1.ListIndex = 0
    CB2.ListIndex = 0
    CB3.ListIndex = 1
    CB4.ListIndex = 1
    CB5.ListIndex = 5
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = True
    cbxQR.Value = False
End Sub

Private Sub B3_Click()
    PlaySound "せんたく"
    CB0.ListIndex = 0
    CB1.ListIndex = 0
    CB2.ListIndex = 0
    CB3.ListIndex = 0
    CB4.ListIndex = 0
    CB5.ListIndex = 0
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 0
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
End Sub

Private Sub B4_Click()
    PlaySound "せんたく"
    CB0.ListIndex = 1
    CB1.ListIndex = 2
    CB2.ListIndex = 0
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 3
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 2
    cbx0.Value = False
    cbx1.Value = True
    cbx2.Value = False
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
End Sub

Private Sub CB0_Change()
    Call CB選択変更
End Sub

Private Sub CB1_Change()
    Call CB選択変更
End Sub

Private Sub CB2_Change()
    Call CB選択変更
End Sub

Private Sub CB3_Change()
    Call CB選択変更
End Sub

Private Sub CB4_Change()
    Call CB選択変更
End Sub

Private Sub CB5_Change()
    If CB5.Value = "" Then Exit Sub
    With ActiveWorkbook.Sheets("製品品番")
        Set key = .Cells.Find("型式", , , 1)
        myCol = .Rows(key.Row).Find(CB5.Value, , , 1).Column
        lastRow = .Cells(.Rows.count, .Cells.Find("メイン品番", , , 1).Column).End(xlUp).Row
        Dim 項目 As String: 項目 = ""
        For i = key.Row + 1 To lastRow
            If InStr(項目, "," & .Cells(i, myCol) & ",") = 0 Then
                項目 = 項目 & "," & .Cells(i, myCol) & ","
            End If
        Next i
    
    End With
    項目 = Mid(項目, 2)
    項目 = Left(項目, Len(項目) - 1)
    項目s = Split(項目, ",,")
    With CB6
        .RowSource = ""
        .Clear
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = -1
    End With
    Call CB選択変更
End Sub

Public Function CB選択変更()
    
    If CB5.Value = "メイン品番" Then 状態 = 0 Else 状態 = 1
    
    状態 = 状態 & CB0.ListIndex & CB1.ListIndex & CB2.ListIndex & CB3.ListIndex & CB4.ListIndex & CB9.ListIndex & CB8.ListIndex
    
    If CB0.ListIndex = 0 Then
        PIC00.Visible = False
    Else
        PIC00.Visible = True
        On Error Resume Next
        PIC00.Picture = LoadPicture(myAddress(0, 1) & "\menu\" & 状態 & ".bmp")
        On Error GoTo 0
    End If
    If サンプル作成モード = True Then
        sample.Visible = True
        sample.Caption = 状態
    End If
End Function

Public Function マジック選択変更()
    Call addressSet(wb(0))
    
    状態 = CB8.Value
    
    状態 = Replace(状態, "-1", "0")
    
    If Left(状態, 1) = "0" Then 状態 = "0000000000"
    
    PIC01.Picture = LoadPicture(myAddress(0, 1) & "\" & 状態 & ".jpg")
End Function

Private Sub CheckBox1_Click()

End Sub

Private Sub CBa_Change()

End Sub

Private Sub CB6_Change()
    Call CB選択変更
End Sub

Private Sub CB8_Change()
    Call CB選択変更
End Sub

Private Sub cbx0_Click()
    If cbx0.Value = True Then
        cbxQR.Visible = True
    Else
        cbxQR.Visible = False
        cbxQR.Value = False
    End If
End Sub

Private Sub cbx1_Change()
    If cbx1.Value = True Then
        CB7.Visible = True
        CB7.ListIndex = 0
    Else
        CB7.Visible = False
        CB7.ListIndex = -1
    End If
End Sub

Private Sub cbx2_Click()
    If cbx2.Value = True Then
        CB9.Visible = True
        Label9.Visible = True
    Else
        CB9.Visible = False
        Label9.Visible = False
    End If
End Sub

Private Sub cbx4_Click()
    If cbx4 = True Then
        If Dir(myAddress(3, 1)) = "" Then
            コメント.Visible = True
            コメント.Caption = "後ハメ作業者取得先が見つかりません。設定を確認してください。"
            コメント.ForeColor = RGB(255, 0, 0)
        End If
    Else
        コメント.Caption = "後ハメ作業者取得先が見つかりました。"
        コメント.ForeColor = 0
        コメント.Visible = False
    End If
End Sub

Private Sub cbx7_Click()

End Sub

Private Sub CommandButton1_Click()
    フォームからの呼び出し = True
    mytime = time
    
    If CB6.ListIndex = -1 Then
        コメント.Visible = True
        コメント.Caption = "作成対象が指定されていません。"
        Beep
        Exit Sub
    End If
    
    If コメント.ForeColor = RGB(255, 0, 0) Then
        MsgBox "Sheet[設定]の後ハメ作業者一覧取得_のアドレスが見つかりません。"
        Exit Sub
    End If
    端末ナンバー表示 = True
    
    PlaySound ("じっこう")
    cb選択 = CB0.ListIndex
    cb選択 = cb選択 & "," & CB1.ListIndex
    cb選択 = cb選択 & "," & CB2.ListIndex
    cb選択 = cb選択 & "," & CB3.ListIndex
    cb選択 = cb選択 & "," & CB4.ListIndex
    cb選択 = cb選択 & "," & CB9.ListIndex
    cb選択 = cb選択 & "," & CB10.ListIndex
    
    マルマ形状 = CB8.List(CB8.ListIndex, 1)
    
    Call 製品品番RAN_set2(製品品番Ran, CB5.Value, CB6.Value, "")
    
    色で判断 = cbx2
    二重係止flg = cbx5
    後ハメ作業者 = cbx4
    QR印刷 = cbxQR
    is_setProcessColor = cbx6
    後ハメ点滅 = cbx7
    
    成型角度無視flag = True
    
    If 製品品番RANc = 0 Then
        コメント.Visible = True
        コメント.Caption = "該当する製品品番がありません。" & vbCrLf _
                         & "例えば選択した条件が、" & vbCrLf & "[PVSW_RLTF]に在りません。"
        Beep
        Exit Sub
    End If
    
    Set myBook = ActiveWorkbook
    
    If 製品品番RANc <> 1 And cbx0.Value = True Then
        コメント.Visible = True
        コメント.Caption = "品番が複数ある為、サブ図作成不可。"
        Beep
        Exit Sub
    End If
    
    Unload UI_01
    
    Call PVSWcsv両端のシート作成_Ver2001
    
    If 後ハメ作業者 = True Then
        If 製品品番RANc = 1 Then
            myAddress(3, 1) = 製品品番Ran(製品品番RAN_read(製品品番Ran, "後ハメ作業者取得"), 1)
            If Dir(myAddress(3, 1)) <> "" Then
                Set wb(3) = Workbooks.Open(fileName:=myAddress(3, 1), UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
                Call SQLもどき_後ハメ作業者(後ハメ作業者ran, CB6.Value)
                Application.DisplayAlerts = False
                wb(3).Close
                Application.DisplayAlerts = True
                '後ハメ点滅の為に後ハメ作業RANにCAVの情報とか入れとく
                Dim tempArray As Variant
                tempArray = readSheetToRan3(wb(0).Sheets("PVSW_RLTF両端"), "電線識別名", CB6.Value & ",RLTFtoPVSW_,端末識別子,キャビティ,構成_,ハメ", "", 1)
                
                For i = LBound(後ハメ作業者ran, 2) + 1 To UBound(後ハメ作業者ran, 2)
                    For ii = LBound(tempArray, 2) To UBound(tempArray, 2)
                        If tempArray(5, ii) = "後" Then
                            If 後ハメ作業者ran(0, i) = tempArray(4, ii) Then
                                後ハメ作業者ran(3, i) = tempArray(2, ii)
                                後ハメ作業者ran(4, i) = tempArray(3, ii)
                                Exit For
                            End If
                        End If
                    Next ii
                Next i
            End If
        Else
            MsgBox "製品品番が１点を超える場合、後ハメ作業者の表示は未だ対応していません。"
            Exit Sub
        End If
    End If
    
    
    'サンプル作成モード
    Dim cb5str(1) As String: cb5str(0) = "メイン品番": cb5str(1) = "結き"
    Dim cb6str(1) As String: cb6str(0) = "8211136Y82     ": cb6str(1) = "G"
    If サンプル作成モード = True Then
        For i0 = 1 To CB0.ListCount - 1
            For i1 = 0 To CB1.ListCount - 1
                For i2 = 0 To CB2.ListCount - 1
                    For i3 = 0 To CB3.ListCount - 1
                        For i4 = 0 To CB4.ListCount - 1
                            For i8 = 0 To CB8.ListCount - 1
                                i9 = -1
                                マルマ形状 = CB8.List(i8, 1)
                                cb選択 = i0 & "," & i1 & "," & i2 & "," & i3 & "," & i4 & "," & i9
                                For ii = 0 To 1
                                    Call 製品品番RAN_set2(製品品番Ran, cb5str(ii), cb6str(ii), "")
                                    Call ハメ図作成_Ver220098(cb選択, cb5str(ii), cb6str(ii))
                                  
                                    On Error Resume Next
                                    ActiveSheet.Shapes.Range("324_1").Select
                                    If err.number = 1004 Then
                                        ActiveSheet.Shapes.Range("324_7").Select
                                    End If
                                    On Error GoTo 0
                                    Call 画像として出力(ii & Replace(cb選択, ",", "") & i8)
                                Next ii
                            Next i8
                        Next i4
                    Next i3
                Next i2
            Next i1
        Next i0
    Else
        Call ハメ図作成_Ver220098(cb選択, CB5.Value, CB6.Value)
    End If
    
    
    
    
    
    If cbx7.Value = True Then Call 後ハメ作業者別_点滅画像作成(CB6.Value, CB12.Value)
    If cbx3.Value = True Then Call 検査履歴システム用データ作成v2182(CB6.Value)
    If cbx1.Value = True Then Call ハメ図の印刷用データ作成(CB7.Value, CB5.Value & Replace(CB6.Value, " ", ""))
    'MsgBox "作成が完了しました。"
    
    If cbx0.Value = True Then
        msg = サブ図作成_Ver220116(CB6.Value, 製品品番Ran)
        If msg <> "" Then
            DoEvents
            a = MsgBox("次のサブが[端末一覧]にあって[PVSW_RLTF]にありません。" & vbCrLf & "マル即等で変更になっていませんか？" & vbCrLf & vbCrLf & msg, , "アンマッチエラー")
        End If
    End If
    
    If マルマ不足 <> "" Then
        マルマ不足sp = Split(マルマ不足, "_")
        msg = "次の端末でマルマの数が不足。マルマ自動立案が正常に終了していません。"
        msg = msg & Join(マルマ不足sp, vbCrLf)
    End If
    
    Call ログ出力("test", "test", "ハメ図" & cb選択 & CB5.Value & CB6.Value)
    
    myBook.Activate
    If サンプル作成モード = True Then
    
    Else
        MsgBox "作成時間= " & DateDiff("s", mytime, time) & " s"
    End If
    
    'Call 完了しました(myBook)
    'mybook.VBProject.VBComponents(Sheets("ハメ図_" & CB5.Value & Replace(CB6.Value, " ", "")).CodeName).CodeModule.AddFromFile myaddress(0,1) & "\002_問連書作成_マルマ.txt"
End Sub

Private Sub OptionButton1_Click()
    
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("もどる")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    PlaySound "せんたく"
    CB0.ListIndex = 4
    CB1.ListIndex = 1
    CB2.ListIndex = 0
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = False
    cbx5.Value = False
    cbxQR.Value = False
End Sub

Private Sub CommandButton6_Click()
    PlaySound "せんたく"
    CB0.ListIndex = 2
    CB1.ListIndex = 0
    CB2.ListIndex = 0
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = -1
    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 0
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = True
    cbx4.Value = False
    cbx5.Value = True
    cbxQR.Value = False
End Sub

Private Sub ST01_Change()

End Sub

Private Sub PIC_Click()

End Sub

Private Sub CommandButton7_Click()
    PlaySound "せんたく"
    CB0.ListIndex = 6
    CB1.ListIndex = 2
    CB2.ListIndex = 0
    CB3.ListIndex = 0
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    CB6.ListIndex = 13

    CB7.ListIndex = -1
    CB8.ListIndex = 0
    CB9.ListIndex = -1
    CB10.ListIndex = 0
    cbx0.Value = False
    cbx1.Value = False
    cbx2.Value = False
    cbx3.Value = False
    cbx4.Value = True
    cbx5.Value = False
    cbx7.Value = True
    cbxQR.Value = False
    
    
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub PIC00_Click()

End Sub

Private Sub UserForm_Initialize()
    Set wb(0) = ThisWorkbook
    Dim 項目(12) As String
    項目(0) = "図を作成しない,電線サイズのみ,ポイント,回路符号,構成,相手端末,後ハメ作業ナンバー"
    項目(1) = "何もしない,先ハメは赤線,先ハメは小さくする,先ハメは塗りつぶす,先ハメのみ表示"
    項目(2) = "表示しない,先ハメ部品(工程40)"
    項目(3) = "変換しない,変換する"
    項目(4) = "使用しない,使用する"
    項目(6) = "A4-タテ,A4-横,A3-タテ,A3-横"
    項目(8) = "Tear,Oval,Heart" 'マルマの形状
    項目(10) = "160,9,21" 'マルマの番号
    項目(11) = "表示しない,後ハメ数=0なら表示,後ハメ数 <> 0なら表示"
    
    項目(12) = "2右"
        
    With ActiveWorkbook.Sheets("設定")
        Set myKey = .Cells.Find("ハメ色_", , , 1)
        lastRow = myKey.Offset(0, 1).End(xlDown).Row - myKey.Row
        For i = 0 To lastRow
            項目(9) = 項目(9) & "," & myKey.Offset(i, 2)
        Next i
        項目(9) = Mid(項目(9), 2)
    End With
    
    With ActiveWorkbook.Sheets("製品品番")
        Set myKey = .Cells.Find("型式", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For x = myKey.Column To lastCol
            If .Cells(myKey.Row, x).Offset(-1, 0) = 1 Then
                項目(5) = 項目(5) & "," & .Cells(myKey.Row, x)
            End If
        Next x
        項目(5) = Mid(項目(5), 2)
        Set myKey = Nothing
    End With
    
    項目s = Split(項目(0), ",")
    With CB0
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = 0
    End With

    項目s = Split(項目(1), ",")
    With CB1
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = 0
    End With
    
    項目s = Split(項目(2), ",")
    With CB2
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = 0
    End With
    
    項目s = Split(項目(3), ",")
    With CB3
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = 0
    End With
    
    項目s = Split(項目(4), ",")
    With CB4
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = 0
    End With
    
    項目s = Split(項目(5), ",")
    With CB5
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = 0
    End With
    
    項目s = Split(項目(6), ",")
    With CB7
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = 0
    End With
    
    項目s = Split(項目(8), ",")
    項目s2 = Split(項目(10), ",")
    With CB8
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem
             .List(i, 0) = 項目s(i)
             .List(i, 1) = 項目s2(i)
        Next i
        .ListIndex = 0
    End With
    
    項目s = Split(項目(9), ",")
    With CB9
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = -1
    End With
    項目s = Split(項目(11), ",")
    With CB10
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = 0
    End With
    
    項目s = Split(項目(12), ",")
    With CB12
        .RowSource = ""
        For i = LBound(項目s) To UBound(項目s)
            .AddItem 項目s(i)
        Next i
        .ListIndex = 0
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "とじる"
End Sub

Private Sub コメント_Click()

End Sub


