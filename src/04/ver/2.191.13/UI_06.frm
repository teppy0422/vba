VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_06 
   Caption         =   "ÆCAEg"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   OleObjectBlob   =   "UI_06.frx":0000
   StartUpPosition =   1  'I[i[ tH[Ì
End
Attribute VB_Name = "UI_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














































































Private Sub CB5_Change()
    
End Sub

Private Sub CB0_Change()
    Dim Ú(1) As String
    Dim Ú2(1) As String
    'CB0.Text
    With ActiveWorkbook.Sheets("»iiÔ")
        Set myKey = .Cells.Find("^®", , , 1)
        Set myKey = .Rows(myKey.Row).Find(CB0.Text, , , 1)
        Set mykey2 = .Rows(myKey.Row).Find("«", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            If InStr(Ú(0), "," & .Cells(Y, myKey.Column)) & "," = 0 Then
                Ú(0) = Ú(0) & "," & .Cells(Y, myKey.Column) & ","
                Ú2(0) = Ú2(0) & "," & .Cells(Y, mykey2.Column) & ","
            End If
        Next Y
        If Len(Ú(0)) <= 2 Then
            Ú(0) = ""
            Ús = Empty
        Else
            Ú(0) = Mid(Ú(0), 2)
            Ú(0) = Left(Ú(0), Len(Ú(0)) - 1)
            Ús = Split(Ú(0), ",,")
            Ú2(0) = Mid(Ú2(0), 2)
            Ú2(0) = Left(Ú2(0), Len(Ú2(0)) - 1)
            Ú2s = Split(Ú2(0), ",,")
        End If
    End With
    
    With CB1
        .RowSource = ""
        .Clear
        If Not IsEmpty(Ús) Then
            For i = LBound(Ús) To UBound(Ús)
                .AddItem
                .List(i, 0) = Ús(i)
                .List(i, 1) = Ú2s(i)
            Next i
            .ListIndex = UBound(Ús)
        End If
    End With
End Sub

Private Sub CB1_Change()
    Call »iiÔRAN_set2(»iiÔRAN, CB0.Value, CB1.Value, "")
    If »iiÔRANc <> 1 Then
'        myLabel.Caption = "»iiÔ_ªÙíÅ·B"
'        myLabel.ForeColor = RGB(255, 0, 0)
        Exit Sub
    Else
        myLabel.Caption = ""
    End If
End Sub

Private Sub CommandButton4_Click()
    PlaySound "àÇé"
    Unload Me
    UI_Menu.Show
End Sub

Private Sub CommandButton5_Click()
    PlaySound "¶Á±¤"
    Call »iiÔRAN_set2(»iiÔRAN, CB0.Value, CB1.Value, "")
    If »iiÔRANc = 0 Then
        Exit Sub
    End If
    Unload Me
    
    Set wb(0) = ActiveWorkbook
    Call OoÍ("test", "test", "ÆCAEg" & CB0.Value & "-" & CB1.Value & "-" & CB2.Value)
    Call ÆCAEg}Ìì¬ver2179(CB0.Value, CB1.Value, CB2.Value)
    PlaySound "©ñ¹¢"
    
End Sub

Private Sub UserForm_Initialize()
    Dim Ú(1) As String
    With ActiveWorkbook.Sheets("»iiÔ")
        Set myKey = .Cells.Find("^®", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For X = myKey.Column To lastCol
            Ú(0) = Ú(0) & "," & .Cells(myKey.Row, X)
        Next X
        Ú(0) = Mid(Ú(0), 2)
    End With
    Ús = Split(Ú(0), ",")
    With CB0
        .RowSource = ""
        For i = LBound(Ús) To UBound(Ús)
            .AddItem Ús(i)
            If Ús(i) = "^®" Then myindex = i
        Next i
        .ListIndex = myindex
    End With

    Call SQL_©®@(©®@RAN)

    With CB2
        .RowSource = ""
        For i = LBound(©®@RAN, 2) To UBound(©®@RAN, 2)
            .AddItem ©®@RAN(0, i)
        Next i
        .ListIndex = -1
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "Æ¶é"
End Sub
