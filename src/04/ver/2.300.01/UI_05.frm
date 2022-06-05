VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_05 
   Caption         =   "その他"
   ClientHeight    =   4650
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7530
   OleObjectBlob   =   "UI_05.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









































































































































































































Private Sub B0_Click()
    PlaySound "せんたく"
    CB0.ListIndex = 4
    CB1.ListIndex = 0
    CB2.ListIndex = 1
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    cbx0.Value = True
    cbx1.Value = False
    cbx2.Value = True
    PIC00.Picture = LoadPicture(myAddress(0, 1) & "\ハメ図sample_" & "4511000000" & ".jpg")
End Sub

Private Sub CB0_Change()
    'Call CB選択変更
End Sub

Private Sub CB5_Change()
    If CB5.Value = "" Then Exit Sub
    With ActiveWorkbook.Sheets("製品品番")
        Set key = .Cells.Find("型式", , , 1)
        myCol = .Rows(key.Row).Find(CB5.Value, , , 1).Column
        lastRow = .Cells(.Rows.count, myCol).End(xlUp).Row
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
End Sub

Public Function CB選択変更()
    Call addressSet(wb(0))
    
    状態 = CB0.ListIndex & CB1.ListIndex & CB2.ListIndex & CB3.ListIndex
    状態 = 状態 & "000000"
    
    状態 = Replace(状態, "-1", "0")
    
    If Left(状態, 1) = "0" Then 状態 = "0000000000"
    
    PIC00.Picture = LoadPicture(myAddress(0, 1) & "\ハメ図sample_" & 状態 & ".jpg")
End Function

Private Sub CommandButton1_Click()
    Unload Me
    If CB6.ListIndex = -1 Then
        コメント.Visible = True
        コメント.Caption = "製品品番が選択されていません。"
        Beep
        Exit Sub
    End If
    
    PlaySound ("じっこう")
    cb選択 = CB0.ListIndex
    
    Call 製品品番RAN_set2(製品品番Ran, CB5.Value, CB6.Value, "")
    
    If 製品品番RANc = 0 Then
        コメント.Visible = True
        コメント.Caption = "該当する製品品番がありません。" & vbCrLf _
                         & "例えば選択した条件が、" & vbCrLf & "[PVSW_RLTF]に在りません。"
        Beep
        Exit Sub
    End If
    
    Unload UI_01
    
    Select Case CB0.ListIndex
    Case 0
        PlaySound ("けってい")
        Call サブ一覧表の作成
        PlaySound ("かんせい")
    Case 1
        PlaySound ("けってい")
        Call 類似コネクタ一覧b作成
        PlaySound ("かんせい")
    Case -1
        
    End Select
    
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("もどる")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub UserForm_Initialize()

    Dim 項目(6) As String
    項目(0) = "サブ一覧表,類似コネクタ一覧b"
    
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

    
    項目s = Split(項目(5), ",")
    With CB5
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

