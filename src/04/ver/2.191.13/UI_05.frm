VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_05 
   Caption         =   "»Ì¼"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   OleObjectBlob   =   "UI_05.frx":0000
   StartUpPosition =   1  'I[i[ tH[Ì
End
Attribute VB_Name = "UI_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














































































































Private Sub B0_Click()
    PlaySound "¹ñ½­"
    CB0.ListIndex = 4
    CB1.ListIndex = 0
    CB2.ListIndex = 1
    CB3.ListIndex = 1
    CB4.ListIndex = 0
    CB5.ListIndex = 1
    cbx0.Value = True
    cbx1.Value = False
    cbx2.Value = True
    PIC00.Picture = LoadPicture(AhX(0) & "\n}sample_" & "4511000000" & ".jpg")
End Sub

Private Sub CB0_Change()
    'Call CBIðÏX
End Sub

Private Sub CB5_Change()
    If CB5.Value = "" Then Exit Sub
    With ActiveWorkbook.Sheets("»iiÔ")
        Set key = .Cells.Find("^®", , , 1)
        myCol = .Rows(key.Row).Find(CB5.Value, , , 1).Column
        lastRow = .Cells(.Rows.count, myCol).End(xlUp).Row
        Dim Ú As String: Ú = ""
        For i = key.Row + 1 To lastRow
            If InStr(Ú, "," & .Cells(i, myCol) & ",") = 0 Then
                Ú = Ú & "," & .Cells(i, myCol) & ","
            End If
        Next i
    
    End With
    Ú = Mid(Ú, 2)
    Ú = Left(Ú, Len(Ú) - 1)
    Ús = Split(Ú, ",,")
    With CB6
        .RowSource = ""
        .Clear
        For i = LBound(Ús) To UBound(Ús)
            .AddItem Ús(i)
        Next i
        .ListIndex = -1
    End With
End Sub

Public Function CBIðÏX()
    Call AhXZbg(myBook)
    
    óÔ = CB0.ListIndex & CB1.ListIndex & CB2.ListIndex & CB3.ListIndex
    óÔ = óÔ & "000000"
    
    óÔ = Replace(óÔ, "-1", "0")
    
    If Left(óÔ, 1) = "0" Then óÔ = "0000000000"
    
    PIC00.Picture = LoadPicture(AhX(0) & "\n}sample_" & óÔ & ".jpg")
End Function

Private Sub CommandButton1_Click()
    Unload Me
    If CB6.ListIndex = -1 Then
        Rg.Visible = True
        Rg.Caption = "»iiÔªIð³êÄ¢Ü¹ñB"
        Beep
        Exit Sub
    End If
    
    PlaySound ("¶Á±¤")
    cbIð = CB0.ListIndex
    
    Call »iiÔRAN_set2(»iiÔRAN, CB5.Value, CB6.Value, "")
    
    If »iiÔRANc = 0 Then
        Rg.Visible = True
        Rg.Caption = "Y·é»iiÔª èÜ¹ñB" & vbCrLf _
                         & "á¦ÎIðµ½ðªA" & vbCrLf & "[PVSW_RLTF]ÉÝèÜ¹ñB"
        Beep
        Exit Sub
    End If
    
    Unload UI_01
    
    Select Case CB0.ListIndex
    Case 0
        PlaySound ("¯ÁÄ¢")
        Call Tuê\Ìì¬
        PlaySound ("©ñ¹¢")
    Case 1
        PlaySound ("¯ÁÄ¢")
        Call ÞRlN^êbì¬
        PlaySound ("©ñ¹¢")
    Case -1
        
    End Select
    
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("àÇé")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub UserForm_Initialize()

    Dim Ú(6) As String
    Ú(0) = "Tuê\,ÞRlN^êb"
    
    With ActiveWorkbook.Sheets("»iiÔ")
        Set myKey = .Cells.Find("^®", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For X = myKey.Column To lastCol
            If .Cells(myKey.Row, X).Offset(-1, 0) = 1 Then
                Ú(5) = Ú(5) & "," & .Cells(myKey.Row, X)
            End If
        Next X
        Ú(5) = Mid(Ú(5), 2)
        Set myKey = Nothing
    End With
    
    Ús = Split(Ú(0), ",")
    With CB0
        .RowSource = ""
        For i = LBound(Ús) To UBound(Ús)
            .AddItem Ús(i)
        Next i
        .ListIndex = 0
    End With

    
    Ús = Split(Ú(5), ",")
    With CB5
        .RowSource = ""
        For i = LBound(Ús) To UBound(Ús)
            .AddItem Ús(i)
        Next i
        .ListIndex = 0
    End With
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "Æ¶é"
End Sub

