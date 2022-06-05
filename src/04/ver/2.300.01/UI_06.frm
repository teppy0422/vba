VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_06 
   Caption         =   "竿レイアウト"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
   OleObjectBlob   =   "UI_06.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









































































































































































Private Sub CB5_Change()
    
End Sub

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
            .ListIndex = UBound(項目s)
        End If
    End With
End Sub

Private Sub CB1_Change()
    Call 製品品番RAN_set2(製品品番Ran, CB0.Value, CB1.Value, "")
    If 製品品番RANc <> 1 Then
'        myLabel.Caption = "製品品番点数が異常です。"
'        myLabel.ForeColor = RGB(255, 0, 0)
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
    PlaySound "じっこう"
    Call 製品品番RAN_set2(製品品番Ran, CB0.Value, CB1.Value, "")
    If 製品品番RANc = 0 Then
        Exit Sub
    End If
    Unload Me
    
    Set wb(0) = ActiveWorkbook
    Call 竿レイアウト図の作成ver2179(CB0.Value, CB1.Value, CB2.Value)
    Call ログ出力("test", "test", "竿レイアウト" & CB0.Value & "-" & CB1.Value & "-" & CB2.Value)
    PlaySound "かんせい"
    
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
            If 項目s(i) = "型式" Then myindex = i
        Next i
        .ListIndex = myindex
    End With

    Call SQL_自動機(自動機RAN)

    With CB2
        .RowSource = ""
        For i = LBound(自動機RAN, 2) To UBound(自動機RAN, 2)
            .AddItem 自動機RAN(0, i)
        Next i
        .ListIndex = -1
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "とじる"
End Sub
