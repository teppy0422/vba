VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_03 
   Caption         =   "zõ}"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   OleObjectBlob   =   "UI_03.frx":0000
   StartUpPosition =   1  'I[i[ tH[Ì
End
Attribute VB_Name = "UI_03"
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
            .ListIndex = 0
        End If
    End With
End Sub

Private Sub CB1_Change()
    Call »iiÔRAN_set2(»iiÔRAN, CB0.Value, CB1.Value, "")
    If »iiÔRANc <> 1 Then
        myLabel.Caption = "»iiÔ_ªÙíÅ·B"
        myLabel.ForeColor = RGB(255, 0, 0)
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
    mytime = Time
    PlaySound "¶Á±¤"
    Call »iiÔRAN_set2(»iiÔRAN, CB0.Value, CB1.Value, "")
    èïstr = »iiÔRAN(»iiÔRAN_read(»iiÔRAN, "«"), 1)
    With myBook.Sheets("èï_" & èïstr)
        .Activate
        Call zõ}ì¬
    End With
    Call »iiÔRAN_set2(»iiÔRAN, CB0.Value, CB1.Value, "") 'zõ}ì¬ÌÉ¯¶èïÌ»iiÔªZbg³êéÌÅZbg
    If »iiÔRANc <> 1 Then
        myLabel.Caption = "»iiÔ_ªÙíÅ·B"
        myLabel.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    Unload Me
    
    Set wb(0) = ActiveWorkbook
    
    cbIð = "1,4,1,1,0,-1"
    }}`ó = 21
    ¶ª = cbx1
    Call n}ì¬_Ver2001(cbIð, CB0.Value, CB1.Value)
    If cbx2 = True Then æn_Å = True
    If cbx1 = True Then
        Call zõ}ì¬one3(»iiÔRAN, "n}_" & CB0.Value & "_" & Replace(CB1.Value, " ", ""))
        Call OoÍ("test", "test", "zõU±" & CB1.Value)
    Else
        Call zõ}ì¬one(»iiÔRAN, "n}_" & CB0.Value & "_" & Replace(CB1.Value, " ", ""))
        Call OoÍ("test", "test", "zõ}" & CB1.Value)
    End If
    Call ÅK»àÇ·
    PlaySound "©ñ¹¢"
    
    Dim myMsg As String: myMsg = "ì¬µÜµ½" & vbCrLf & DateDiff("s", mytime, Time) & "s"
    If zõ}ì¬temp = 1 Then myMsg = myMsg & vbCrLf & vbCrLf & "¡ïÀWf[^ª©Â©çÈ©Á½ÌÅãn}f[^ÌÝì¬µÜµ½B"
    aa = MsgBox(myMsg, vbOKOnly, "¶Yõ+zõU±")
End Sub

Private Sub CommandButton6_Click()
    
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
            If Ús(i) = "CiÔ" Then myindex = i
        Next i
        .ListIndex = myindex
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "Æ¶é"
End Sub
