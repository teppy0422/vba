VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_10 
   Caption         =   "Tuio[ÌoÍ"
   ClientHeight    =   4485
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4320
   OleObjectBlob   =   "UI_10.frx":0000
   StartUpPosition =   1  'I[i[ tH[Ì
End
Attribute VB_Name = "UI_10"
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
        Label_alert.Caption = "»iiÔ_ªÙíÅ·B"
        Label_alert.ForeColor = RGB(255, 0, 0)
        Exit Sub
    Else
        Label_alert.Caption = ""
    End If
End Sub

Private Sub CommandButton1_Click()
    Call addressSet(wb(0))
    
    If myAddress(2, 1) = "" Then
        Call MsgBox("±ÌIPÅÍt@Cªo^³êÄ¢Ü¹ñB", vbOKOnly, "Sjp+")
        Exit Sub
    End If
    
    If Label_alert.ForeColor = 255 Then MsgBox "ÝèðmFµÄ­¾³¢", , "ÀsÅ«Ü¹ñ": Exit Sub
    PlaySound ("¯ÁÄ¢")
    
    Call »iiÔRAN_set2(»iiÔRan, CB0.Value, CB1.Value, "")
    
    Dim ÝÏstr As String
    
    »iiÔstr = UI_10.CB1.Value
    ÝÏstr = »iiÔRan(»iiÔRAN_read(»iiÔRan, "èz"), 1)
    Dim myMessage As String
    If CheckBox_wireEfu Then
        myMessage = myMessage & _
                               "düÌTuðXVµÜ·B" & vbCrLf & vbCrLf & _
                               "    f[^³: ±ÌubNÌV[g[PVSW_RLTF]" & vbCrLf & _
                               "    oÍæF" & myAddress(2, 1) & vbCrLf & vbCrLf & _
                               "    »¢w¦óüVXeÅt^·éTuÅ·B" & vbCrLf & vbCrLf
    End If
    If CheckBox_tubeEfu Then
        myMessage = myMessage & _
                               "`[uÌTuÆ[ðXVµÜ·B" & vbCrLf & vbCrLf & _
                               "f[^³: ±ÌubNÌV[g[iXg,[ê]" & vbCrLf & _
                               "oÍæF" & myAddress(3, 1) & vbCrLf & vbCrLf & _
                               "`[uGtóüSYSÅt^·éTuÅ·B" & vbCrLf & vbCrLf
    End If
    If CheckBox_partsEfu Then
        myMessage = myMessage & _
                               "p[cÌTuÆ[ðXVµÜ·B" & vbCrLf & vbCrLf & _
                               "f[^³: [[ê]=Tu,[iXg]=»êÈO" & vbCrLf & _
                               "oÍæF±ÌubNÌ[iGt](bè)" & vbCrLf & vbCrLf & _
                               "âiiÇÅt^·éTuÅ·B" & vbCrLf & vbCrLf
    End If
    
    If myMessage = "" Then Exit Sub
    Dim a As Long
    a = MsgBox(myMessage, vbYesNo, "Tuio[XV")
    If a = 6 Then
        Unload Me
        If CheckBox_wireEfu Then Call PVSWcsv©çGtóüpTuio[txtoÍ_Ver2012(myIP, CheckBox_stepNumberAdd)
        If CheckBox_tubeEfu Then Call export_tubeEfu(myIP)
        If CheckBox_partsEfu Then Call export_partEfu(»iiÔstr, ÝÏstr)
        MsgBox "oÍµÜµ½"
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
    
    Call checkSheet("PVSW_RLTF;[ê", wb(0), True, True)
    
    '[ê©çgp·éTuio[ðQbg
    With wb(0).Sheets("[ê")
        Dim myKey As Variant, i As Long, [ As String, Turan() As Variant, foundFlag As Boolean, Tu As String
        ReDim Turan(0, 0)
        Set myKey = .Cells.Find(»iiÔRan(1, 1), , , 1)
        For i = myKey.Row + 1 To .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            Tu = .Cells(i, myKey.Column)
            foundFlag = False
            If Tu <> "" Then
                For x = LBound(Turan, 2) To UBound(Turan, 2)
                    If Turan(0, x) = Tu Then
                        foundFlag = True
                        Exit For
                    End If
                Next x
                If foundFlag = False Then
                    ReDim Preserve Turan(0, UBound(Turan, 2) + 1)
                    Turan(0, UBound(Turan, 2)) = Tu
                End If
            End If
        Next i
        If UBound(Turan, 2) = 0 Then
            MsgBox "[[ê]ÉTuio[ª èÜ¹ñB"
            Stop
        End If
        Turan = WorksheetFunction.transpose(Turan) 'bubbleSort2Ì×ÉüêÖ¦é
        Call BubbleSort2(Turan, 1)
        Turan = WorksheetFunction.transpose(Turan) 'bubbleSort2Ì×ÉüêÖ¦é
    End With
    
    '[ê©ç[ÌTuio[ðQbg
    With wb(0).Sheets("[ê")
        Dim [TuRAN()
        ReDim [TuRAN(1, 0)
        Dim [Col As Long: [Col = .Cells.Find("[", , , 1).Column
        For i = myKey.Row + 1 To .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            Tu = .Cells(i, myKey.Column)
            [ = .Cells(i, [Col)
            If Tu <> "" Then
                ReDim Preserve [TuRAN(1, UBound([TuRAN, 2) + 1)
                [TuRAN(0, UBound([TuRAN, 2)) = Tu
                [TuRAN(1, UBound([TuRAN, 2)) = [
            End If
        Next i
    End With
    
    'PVSW_RLTF©çððQbg
    Set myKey = ws(0).Cells.Find(»iiÔRan(1, 1), , , 1)
    'gp·étB[h¼ÌZbg
    Dim fieldname As String: fieldname = myKey.Value & ",RLTFtoPVSW_,n_¤[¯Êq,I_¤[¯Êq,n_¤LreB,I_¤LreB,Ú±G_,¼[n,\¬_"
    ff = Split(fieldname, ",")
    ReDim f(UBound(ff))
    For x = LBound(ff) To UBound(ff)
        f(x) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(x), , , 1).Column
    Next x
    a = UBound(ff) + 2
    
    Dim lastRow As Long
    lastRow = ws(0).Cells(ws(0).Rows.count, myKey.Column).End(xlUp).Row
    
    'Tuio[ÉdüððZbgµÄ¢­
    Dim myRan() As Variant, y As Long, Tustr As String, r As Long
    ReDim myRan(a, 0)
    For y = LBound(Turan) + 1 To UBound(Turan)
        For x = 0 To 1
            For i = myKey.Row + 1 To lastRow
                Tustr = ws(0).Cells(i, myKey.Column).Value
                ¼[n = ws(0).Cells(i, f(7)).Value
                If Turan(y) = Tustr Then
                    If ¼[n = CStr(x) Then
                        ReDim Preserve myRan(a, UBound(myRan, 2) + 1)
                        For r = LBound(myRan) To UBound(myRan) - 2
                            myRan(r, UBound(myRan, 2)) = ws(0).Cells(i, f(r)).Value
                        Next r
                    End If
                End If
            Next i
        Next x
    Next y
    
    Call ÅK»àÇ·
    PlaySound "©ñ¹¢"
    
    Dim myMsg As String: myMsg = "ì¬µÜµ½" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "ænU±_SSC³µ")
End Sub

Private Sub myLabel_Click()
    
End Sub

Private Sub Label8_Click()

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
