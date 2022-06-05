VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_70 
   Caption         =   "70_検査履歴"
   ClientHeight    =   3000
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4800
   OleObjectBlob   =   "UI_70.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_70"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False































































Private Sub cbx1_Click()
    If cbx1 = True Then

    Else
        Label0.Visible = True
        Label0.Caption = ""
    End If
End Sub

Private Sub CommandButton1_Click()

    Call 製品品番RAN_set2(製品品番RAN, "メイン品番", CB0.Value, "")
    マルマ形状 = 160 'Tear
    If cbx0 = True Then
        cb選択 = "2,0,0,1,0,-1"
        色で判断 = True
    Else
        cb選択 = "2,1,0,1,0,-1"
        色で判断 = False
    End If
    Unload Me
    Call ハメ図作成_Ver2001(cb選択, "メイン品番", CB0.Value)
    Call 検査履歴システム用データ作成v2182(CB0.Value)
    myBook.Activate
    MsgBox "作成しました" & vbLf & DateDiff("s", mytime, Time) & " s"
End Sub

Private Sub CommandButton4_Click()
    PlaySound "もどる"
    Unload Me
    UI_Menu.Show
End Sub

Private Sub UserForm_Initialize()
    Dim 項目(0) As String
    Call アドレスセット(myBook)
    With ActiveWorkbook.Sheets("製品品番")
        Set myKey = .Cells.Find("メイン品番", , , 1)
        lastRow = .Cells(Rows.count, myKey.Column).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            項目(0) = 項目(0) & "," & .Cells(Y, myKey.Column)
        Next Y
        項目(0) = Mid(項目(0), 2)
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
End Sub

