VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_11 
   Caption         =   "i \¦_ì¬"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
   OleObjectBlob   =   "UI_11.frx":0000
   StartUpPosition =   1  'I[i[ tH[Ì
End
Attribute VB_Name = "UI_11"
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
    
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    mytime = time
    PlaySound "¶Á±¤"
    Call »iiÔRAN_set2(»iiÔRan, CB0.Value, CB1.Value, "")
    
    Dim fileName As String: fileName = Replace(wb(0).Name, ".xlsm", "") & "_i \¦_" & CB0.Value & "_" & CB1.Value & ".xlsx"
    Unload Me
    
    Dim i As Long, pNumbers As String
    For i = LBound(»iiÔRan, 2) + 1 To UBound(»iiÔRan, 2)
        pNumbers = pNumbers & "," & »iiÔRan(»iiÔRAN_read(»iiÔRan, "CiÔ"), i)
    Next i
    
    Dim setWords As String, setWordsSP As Variant
    setWords = "iiÔ,ÄÌ,d,D,W,L,F,i¼Ì,ÞÚ×,íÞ,Hö,Höa"
    setWordsSP = Split(setWords, ",")
    
    Set ws(1) = wb(0).Sheets("iXg")
    msg = checkFieldName("iiÔ", ws(1), setWords)
    If msg <> "" Then
        msg = "[iXg]ÉÌtB[hª©Â©èÜ¹ñB" & msg & vbCrLf & vbCrLf & _
                   "±Ì@\ðgp·éÉÍVer2.200.70È~Åì¬µ½[iXg]Å éKvª èÜ·B" & vbCrLf & _
                   "ì¬ð~µÜ·B"
        MsgBox msg, vbOKOnly, "PLUS+"
        End
    End If
    
    Dim Array_iXg As Variant
    Array_iXg = readSheetToRan2(ws(1), "iiÔ", setWords & pNumbers, "")
    
    '»iiÔÉgpª³¢iðí
    Dim x As Long, skipFlag As Boolean
    For i = LBound(Array_iXg, 2) + 1 To UBound(Array_iXg, 2)
        skipFlag = True
        For x = UBound(setWordsSP) + 1 To UBound(setWordsSP) + UBound(»iiÔRan, 2)
            If Array_iXg(x, i) <> "" Then
                skipFlag = False
                Exit For
            End If
        Next x
        If skipFlag = True Then
            Debug.Print i, Array_iXg(0, i)
            Array_iXg = delete_RanVer2(Array_iXg, i)
            i = i - 1
        End If
        If i + 1 > UBound(Array_iXg, 2) Then Exit For
    Next i
    
    'oÍ·éf[^ÌÜÆß
    Dim addArray() As Variant
    ReDim addArray(2, UBound(Array_iXg, 2))
    addArray(0, 0) = "A"
    addArray(1, 0) = "B"
    addArray(2, 0) = "C"
    
    For i = LBound(Array_iXg, 2) + 1 To UBound(Array_iXg, 2)
        If Array_iXg(9, i) = "B" Then
            addArray(0, i) = Array_iXg(0, i)
            If Array_iXg(2, i) <> "" Then
                addArray(1, i) = Replace(Array_iXg(2, i) & " L=" & Array_iXg(5, i), ".0", "")
            End If
        ElseIf Array_iXg(9, i) = "T" Then
            addArray(2, i) = Replace(Array_iXg(0, i), "-", " ")
            addArray(1, i) = Replace(Array_iXg(1, i), " ", "") & "-" & Array_iXg(6, i)
            If Replace(Array_iXg(3, i), " ", "") <> "" Then
                addArray(0, i) = "D" & Replace(Replace(Array_iXg(2, i) & "~" & Array_iXg(3, i), ".0", ""), " ", "") & " L=" & Replace(Array_iXg(5, i), " ", "")
            ElseIf Replace(Array_iXg(2, i), " ", "") <> "" Then
                addArray(0, i) = "D" & Replace(Replace(Array_iXg(2, i), ".0", ""), " ", "") & " L=" & Replace(Replace(Array_iXg(5, i), ".0", ""), " ", "")
            ElseIf Replace(Array_iXg(4, i), " ", "") <> "" Then
                addArray(0, i) = Replace(Replace("W" & Array_iXg(4, i), ".0", ""), " ", "") & " L=" & Replace(Replace(Array_iXg(5, i), ".0", ""), " ", "")
            End If
        Else
            
        End If
    Next i
    
    '·³Lð0ß4É·é
    Dim array_temp
    For i = LBound(Array_iXg, 2) To UBound(Array_iXg, 2)
        array_temp = Array_iXg(5, i)
        If array_temp <> "" Then
            If IsNumeric(array_temp) Then
                array_temp = Int(array_temp)
                If (Len(array_temp) <= 4) Then
                    array_temp = String(4 - Len(array_temp), "0") & array_temp
                    Array_iXg(5, i) = array_temp
                    
                End If
            End If
        End If
    Next
     
    Array_iXg = merge_Array(addArray, Array_iXg)

    export_ArrayToSheet Array_iXg, "i \¦", True
    
    Dim outputDirectory As String
    outputDirectory = wb(0).path & "\42_i \¦"
    If Dir(outputDirectory, vbDirectory) = "" Then MkDir outputDirectory
    
    'evÌSPC100ÌcsvÆtxtÍAJ}ðæØè¶ÆµÄF¯·é×AJ}æØèÅoÍ·éÆeLXgàÌJ}ÅñYª­¶·é _
    ÈÌÅxlsxÅoÍ
    wb(0).Sheets("i \¦").Move
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=outputDirectory & "\" & fileName, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    Call ÅK»àÇ·
    PlaySound "©ñ¹¢"
    
    Dim myMsg As String: myMsg = "ì¬µÜµ½" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "i \¦_ì¬")
End Sub

Private Sub UserForm_Initialize()
    Dim Ú(1) As String
    With wb(0).Sheets("»iiÔ")
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
            If Ús(i) = "«" Then myindex = i
        Next i
        .ListIndex = myindex
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "Æ¶é"
End Sub
