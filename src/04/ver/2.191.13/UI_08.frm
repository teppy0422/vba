VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_08 
   Caption         =   "Tu§Ä"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   OleObjectBlob   =   "UI_08.frx":0000
   StartUpPosition =   1  'I[i[ tH[Ì
End
Attribute VB_Name = "UI_08"
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

    Unload Me
    Set wb(0) = ActiveWorkbook
    
    'gp·étB[h¼ÌZbg
    Dim fieldName As String: fieldName = "RLTFtoPVSW_,n_¤[¯Êq,I_¤[¯Êq,n_¤[îèiÔ,I_¤[îèiÔ,dã¡@_"
    ff = Split(fieldName, ",")
    Dim f As Variant: ReDim f(UBound(ff))
    For X = LBound(ff) To UBound(ff)
        f(X) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(X), , , 1).Column
    Next X
    a = UBound(ff) + 1
    'düðZbg·ézñ
    Dim [düRAN As Variant
    ReDim [düRAN(a, 0)
    'tB[h¼ðzñÉüêé
    For X = LBound(ff) To UBound(ff)
        [düRAN(X, 0) = ff(X)
    Next X
    'ÎÛÌO[vÉ
    Dim CiÔi As Integer
    CiÔi = »iiÔRAN_read(»iiÔRAN, "CiÔ")
    For i = LBound(»iiÔRAN, 2) + 1 To UBound(»iiÔRAN, 2)
        »iiÔstr = »iiÔRAN(CiÔi, i)
        With wb(0).Sheets("PVSW_RLTF")
            '»iiÔÌtB[hðL[ÆµÄZbg
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(»iiÔstr, , , 1)
            Dim lastRow As Long
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For Y = myKey.Row + 1 To lastRow
                If .Cells(Y, myKey.Column) <> "" Then
                    If .Cells(Y, f(0)) = "Found" Then
                        ReDim Preserve [düRAN(a, UBound([düRAN, 2) + 1)
                        For X = LBound(f) + 1 To UBound(f)
                            [düRAN(X, UBound([düRAN, 2)) = .Cells(Y, f(X))
                        Next X
                        '[düRAN(0, UBound([düRAN, 2)) = 1
                    End If
                End If
            Next Y
        End With
       Call ReplaceLR([düRAN)
       Call SumRan([düRAN)
       'ñHÌ½³©çeðßé
       Dim []¿RAN()
       []¿RAN = evaluationRan([düRAN) 'Dæª9ÌTuio[999
       []¿RAN = changeRowCol([]¿RAN)
       Call BubbleSort3([]¿RAN, 3, 2)
       [düRAN = changeRowCol([düRAN)
    Next i
    
    '[Tuio[999ÌdüTuio[ð999É·é
    For i = LBound([]¿RAN) To UBound([]¿RAN)
        [str = []¿RAN(i, 0)
        Tustr = []¿RAN(i, 5)
        If Tustr <> "" Then
            For ii = LBound([düRAN) To UBound([düRAN)
                For X = 1 To 2
                    If [str = [düRAN(ii, X) Then
                        [düRAN(ii, 6) = Tustr
                        Exit For
                    End If
                Next X
            Next ii
        End If
    Next i
    
    ']¿Ì¢eðîÉTuio[ðzzµÄ¢­
    For i = LBound([]¿RAN) + 1 To UBound([]¿RAN)
        [str = []¿RAN(i, 0)
        è[str = []¿RAN(i, 6)
        'If [str = "250" Then Stop
        If []¿RAN(i, 5) <> "" Then GoTo line20
        []¿RAN(i, 5) = [str
        For j = LBound([düRAN) + 1 To UBound([düRAN)
            If [düRAN(j, 6) = "" Then 'Ü¾Tuio[ªÜÁÄ³¯êÎ
                For X = 1 To 2
                    If [str = [düRAN(j, X) Then
                        []¿lng = [düRAN(j, 0)
                        If X = 1 Then è[str = [düRAN(j, 2)
                        If X = 2 Then è[str = [düRAN(j, 1)
                        'àµè[ª1ÌêÉè[ÌTuio[ÉÏX
                        If è[str = "1" Then
                            è[Tustr = search[]¿RAN([]¿RAN, è[str, 5)
                            If è[Tustr <> "" Then
                                []¿RAN(i, 5) = è[Tustr
                                GoTo line20
                            End If
                        End If
                        è[Dæ = search[]¿RAN([]¿RAN, è[str, 3)
                        If è[Dæ = "1" Then GoTo line15
                        è[]¿lng = searchè[]¿([düRAN, è[str)
                        If []¿lng >= è[]¿lng Then
                            [düRAN(j, 6) = [str
                            For ii = LBound([]¿RAN) + 1 To UBound([]¿RAN)
                                If []¿RAN(ii, 0) = è[str Then
                                    If []¿RAN(ii, 5) = "" Then
                                        []¿RAN(ii, 5) = [str
                                    End If
                                    Exit For
                                End If
                            Next ii
                        End If
                    End If
line15:
                Next X
            End If
        Next j
line20:
    Next i
    
    'qªçÈ©Á½düð]¿Ì¢[É}·æ¤É·é
    For i = LBound([düRAN) + 1 To UBound([düRAN)
        If [düRAN(i, 0) <> "" Then
            If [düRAN(i, 6) = "" Then
                [1str = [düRAN(i, 1)
                [2str = [düRAN(i, 2)
                []¿1str = search[]¿RAN([]¿RAN, [1str, 2)
                []¿2str = search[]¿RAN([]¿RAN, [2str, 2)
                If []¿1str > []¿2str Then
                    ®Tustr = search[]¿RAN([]¿RAN, [1str, 5)
                Else
                    ®Tustr = search[]¿RAN([]¿RAN, [2str, 5)
                End If
                [düRAN(i, 6) = ®Tustr
            End If
        End If
    Next i
    
    addRow = 1
    For Y = LBound([düRAN) To UBound([düRAN)
        If [düRAN(Y, 0) <> "" Then
            For X = LBound([düRAN, 2) To UBound([düRAN, 2)
                With Sheets("temp")
                    .Cells(addRow, X + 1) = [düRAN(Y, X)
                End With
            Next X
            addRow = addRow + 1
        End If
    Next Y
    
    addRow = 1
    For Y = LBound([]¿RAN) To UBound([]¿RAN)
            For X = LBound([]¿RAN, 2) To UBound([]¿RAN, 2)
                With Sheets("temp2")
                    .Cells(addRow, X + 1) = []¿RAN(Y, X)
                End With
            Next X
            addRow = addRow + 1
    Next Y
    Stop
    
    'PVSW_RLTFÉTuio[ðzz
    For i = LBound(»iiÔRAN, 2) + 1 To UBound(»iiÔRAN, 2)
        »iiÔstr = »iiÔRAN(CiÔi, i)
        With wb(0).Sheets("PVSW_RLTF")
            '»iiÔÌtB[hðL[ÆµÄZbg
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(»iiÔstr, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For Y = myKey.Row + 1 To lastRow
                If .Cells(Y, myKey.Column) <> "" Then
                    If .Cells(Y, f(0)) = "Found" Then
                        [str1 = .Cells(Y, f(1))
                        [str2 = .Cells(Y, f(2))
                        '[str1ð¬³¢Éµ¦é
                        swapflg = False
                        If [str1 = "" Then swapflg = True
                        If IsNumeric([str1) = True And IsNumeric([str2) = True Then
                            If Val([str1) > Val([str2) Then
                                swapflg = True
                            End If
                        End If
                        If swapflg = True Then
                            vSwap = [str2
                            [str2 = [str1
                            [str1 = vSwap
                        End If
                        If [str1 & [str2 <> "" Then
                            Tustr = search[düRAN([düRAN, [str1, [str2, 6)
                            If Tustr = "" Then Stop 'SÒÉA
                            .Cells(Y, myKey.Column) = Tustr
                            .Cells(Y, myKey.Column).Interior.color = RGB(129, 216, 208)
                        End If
                    End If
                End If
            Next Y
        End With
    Next i
    
    '[êÉTuio[ðzz
    For i = LBound(»iiÔRAN, 2) + 1 To UBound(»iiÔRAN, 2)
        »iiÔstr = »iiÔRAN(CiÔi, i)
        Set ws(3) = wb(0).Sheets("[ê")
        With ws(3)
            Dim myCol(1) As Integer
            myCol(0) = .Cells.Find("[îèiÔ", , , 1).Column
            myCol(1) = .Cells.Find("[", , , 1).Column
            '»iiÔÌtB[hðL[ÆµÄZbg
            Set myKey = ws(3).Cells.Find(»iiÔstr, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For Y = myKey.Row + 1 To lastRow
                If .Cells(Y, myKey.Column) <> "" Then
                    [îèiÔstr = .Cells(Y, myCol(0))
                    [str = .Cells(Y, myCol(1))
                    Tustr = search[]¿RAN_2pos([]¿RAN, [str, [îèiÔstr, 5)
                    If Tustr = "" Then Stop
                    .Cells(Y, myKey.Column) = Tustr
                    .Cells(Y, myKey.Column).Interior.color = RGB(129, 216, 208)
                End If
            Next Y
        End With
    Next i
    
    Stop
    
    Call ÅK»àÇ·
    PlaySound "©ñ¹¢"
    
    Dim myMsg As String: myMsg = "ì¬µÜµ½" & vbCrLf & DateDiff("s", mytime, Time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "¶Yõ+zõU±")
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
