VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_09 
   Caption         =   "ænU±_SSC³µ"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
   OleObjectBlob   =   "UI_09.frx":0000
   StartUpPosition =   1  'I[i[ tH[Ì
End
Attribute VB_Name = "UI_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
















































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
        For y = myKey.Row + 1 To lastRow
            If InStr(Ú(0), "," & .Cells(y, myKey.Column)) & "," = 0 Then
                Ú(0) = Ú(0) & "," & .Cells(y, myKey.Column) & ","
                Ú2(0) = Ú2(0) & "," & .Cells(y, mykey2.Column) & ","
            End If
        Next y
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
    Call »iiÔRAN_set2(»iiÔRan, CB0.Value, CB1.Value, "")
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
    'í·é
    Set wb(0) = ThisWorkbook
    
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    mytime = time
    PlaySound "¶Á±¤"
    Call »iiÔRAN_set2(»iiÔRan, CB0.Value, CB1.Value, "")

    Unload Me
    
    'n}ðì¬
    cbIð = "4,4,1,1,0,-1,1"
    }}`ó = 160
    ¬^px³flag = True
    [io[\¦ = False
    Call n}ì¬_Ver220098(cbIð, "CiÔ", CB1.Value)
    
    »iiÔstr = Replace(»iiÔRan(»iiÔRAN_read(»iiÔRan, "CiÔ"), 1), " ", "")
    ÝÏstr = »iiÔRan(»iiÔRAN_read(»iiÔRan, "èz"), 1)
    
    'ðÌZbg
    Dim myRan As Variant, myPath As String
    'myRan = setWorkRan([TuRAN)
    myPath = wb(0).path & dirString_09 & Replace(»iiÔstr, " ", "") & "_wire.txt"
    myRan = readTextToArray(myPath)
    
    Call ænU±_SSC³µ(myRan, "n}_CiÔ_" & »iiÔstr, »iiÔstr & "_" & ÝÏstr, [TuRAN)
    
    PlaySound "©ñ¹¢"
    
    Dim myMsg As String: myMsg = "ì¬µÜµ½" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "ænU±_SSC³µ")
End Sub

Private Sub UserForm_Initialize()
    Dim Ú(1) As String
    With ActiveWorkbook.Sheets("»iiÔ")
        Set myKey = .Cells.Find("^®", , , 1)
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        For x = myKey.Column To lastCol
            Ú(0) = Ú(0) & "," & .Cells(myKey.Row, x)
        Next x
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
