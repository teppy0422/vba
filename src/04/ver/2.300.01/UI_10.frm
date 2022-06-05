VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_10 
   Caption         =   "サブナンバーの出力"
   ClientHeight    =   4485
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4320
   OleObjectBlob   =   "UI_10.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_10"
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
        Label_alert.Caption = "製品品番点数が異常です。"
        Label_alert.ForeColor = RGB(255, 0, 0)
        Exit Sub
    Else
        Label_alert.Caption = ""
    End If
End Sub

Private Sub CommandButton1_Click()
    Call addressSet(wb(0))
    
    If myAddress(2, 1) = "" Then
        Call MsgBox("このIPではファイルが登録されていません。", vbOKOnly, "Sjp+")
        Exit Sub
    End If
    
    If Label_alert.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    
    Call 製品品番RAN_set2(製品品番Ran, CB0.Value, CB1.Value, "")
    
    Dim 設変str As String
    
    製品品番str = UI_10.CB1.Value
    設変str = 製品品番Ran(製品品番RAN_read(製品品番Ran, "手配"), 1)
    Dim myMessage As String
    If CheckBox_wireEfu Then
        myMessage = myMessage & _
                               "電線のサブ№を更新します。" & vbCrLf & vbCrLf & _
                               "    データ元: このブックのシート[PVSW_RLTF]" & vbCrLf & _
                               "    出力先：" & myAddress(2, 1) & vbCrLf & vbCrLf & _
                               "    製造指示書印刷システムで付与するサブ№です。" & vbCrLf & vbCrLf
    End If
    If CheckBox_tubeEfu Then
        myMessage = myMessage & _
                               "チューブのサブ№と端末№を更新します。" & vbCrLf & vbCrLf & _
                               "データ元: このブックのシート[部品リスト,端末一覧]" & vbCrLf & _
                               "出力先：" & myAddress(3, 1) & vbCrLf & vbCrLf & _
                               "チューブエフ印刷SYSで付与するサブ№です。" & vbCrLf & vbCrLf
    End If
    If CheckBox_partsEfu Then
        myMessage = myMessage & _
                               "パーツのサブ№と端末№を更新します。" & vbCrLf & vbCrLf & _
                               "データ元: [端末一覧]=サブ№,[部品リスト]=それ以外" & vbCrLf & _
                               "出力先：このブックの[部品エフ](暫定)" & vbCrLf & vbCrLf & _
                               "補給品部品管理で付与するサブ№です。" & vbCrLf & vbCrLf
    End If
    
    If myMessage = "" Then Exit Sub
    Dim a As Long
    a = MsgBox(myMessage, vbYesNo, "サブナンバー更新")
    If a = 6 Then
        Unload Me
        If CheckBox_wireEfu Then Call PVSWcsvからエフ印刷用サブナンバーtxt出力_Ver2012(myIP, CheckBox_stepNumberAdd)
        If CheckBox_tubeEfu Then Call export_tubeEfu(myIP)
        If CheckBox_partsEfu Then Call export_partEfu(製品品番str, 設変str)
        MsgBox "出力しました"
    End If
End Sub

Private Sub CommandButton4_Click()
    PlaySound "もどる"
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    '削除する
    Set wb(0) = ThisWorkbook
    
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    mytime = time
    PlaySound "じっこう"
    Call 製品品番RAN_set2(製品品番Ran, CB0.Value, CB1.Value, "")
    
    Unload Me
    
    Call checkSheet("PVSW_RLTF;端末一覧", wb(0), True, True)
    
    '端末一覧から使用するサブナンバーをゲット
    With wb(0).Sheets("端末一覧")
        Dim myKey As Variant, i As Long, 端末 As String, サブran() As Variant, foundFlag As Boolean, サブ As String
        ReDim サブran(0, 0)
        Set myKey = .Cells.Find(製品品番Ran(1, 1), , , 1)
        For i = myKey.Row + 1 To .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            サブ = .Cells(i, myKey.Column)
            foundFlag = False
            If サブ <> "" Then
                For x = LBound(サブran, 2) To UBound(サブran, 2)
                    If サブran(0, x) = サブ Then
                        foundFlag = True
                        Exit For
                    End If
                Next x
                If foundFlag = False Then
                    ReDim Preserve サブran(0, UBound(サブran, 2) + 1)
                    サブran(0, UBound(サブran, 2)) = サブ
                End If
            End If
        Next i
        If UBound(サブran, 2) = 0 Then
            MsgBox "[端末一覧]にサブナンバーがありません。"
            Stop
        End If
        サブran = WorksheetFunction.transpose(サブran) 'bubbleSort2の為に入れ替える
        Call BubbleSort2(サブran, 1)
        サブran = WorksheetFunction.transpose(サブran) 'bubbleSort2の為に入れ替える
    End With
    
    '端末一覧から端末№毎のサブナンバーをゲット
    With wb(0).Sheets("端末一覧")
        Dim 端末サブRAN()
        ReDim 端末サブRAN(1, 0)
        Dim 端末Col As Long: 端末Col = .Cells.Find("端末№", , , 1).Column
        For i = myKey.Row + 1 To .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            サブ = .Cells(i, myKey.Column)
            端末 = .Cells(i, 端末Col)
            If サブ <> "" Then
                ReDim Preserve 端末サブRAN(1, UBound(端末サブRAN, 2) + 1)
                端末サブRAN(0, UBound(端末サブRAN, 2)) = サブ
                端末サブRAN(1, UBound(端末サブRAN, 2)) = 端末
            End If
        Next i
    End With
    
    'PVSW_RLTFから条件をゲット
    Set myKey = ws(0).Cells.Find(製品品番Ran(1, 1), , , 1)
    '使用するフィールド名のセット
    Dim fieldname As String: fieldname = myKey.Value & ",RLTFtoPVSW_,始点側端末識別子,終点側端末識別子,始点側キャビティ,終点側キャビティ,接続G_,両端ハメ,構成_"
    ff = Split(fieldname, ",")
    ReDim f(UBound(ff))
    For x = LBound(ff) To UBound(ff)
        f(x) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(x), , , 1).Column
    Next x
    a = UBound(ff) + 2
    
    Dim lastRow As Long
    lastRow = ws(0).Cells(ws(0).Rows.count, myKey.Column).End(xlUp).Row
    
    'サブナンバー順に電線条件をセットしていく
    Dim myRan() As Variant, y As Long, サブstr As String, r As Long
    ReDim myRan(a, 0)
    For y = LBound(サブran) + 1 To UBound(サブran)
        For x = 0 To 1
            For i = myKey.Row + 1 To lastRow
                サブstr = ws(0).Cells(i, myKey.Column).Value
                両端ハメ = ws(0).Cells(i, f(7)).Value
                If サブran(y) = サブstr Then
                    If 両端ハメ = CStr(x) Then
                        ReDim Preserve myRan(a, UBound(myRan, 2) + 1)
                        For r = LBound(myRan) To UBound(myRan) - 2
                            myRan(r, UBound(myRan, 2)) = ws(0).Cells(i, f(r)).Value
                        Next r
                    End If
                End If
            Next i
        Next x
    Next y
    
    Call 最適化もどす
    PlaySound "かんせい"
    
    Dim myMsg As String: myMsg = "作成しました" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "先ハメ誘導_SSC無し")
End Sub

Private Sub myLabel_Click()
    
End Sub

Private Sub Label8_Click()

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
