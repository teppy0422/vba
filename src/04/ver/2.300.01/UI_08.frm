VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_08 
   Caption         =   "Tu§Ä"
   ClientHeight    =   3330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5110
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
    mytime = time
    PlaySound "¶Á±¤"
    Call »iiÔRAN_set2(»iiÔRan, CB0.Value, CB1.Value, "")

    Unload Me
    Call checkSheet("PVSW_RLTF;[ê", wb(0), True, True)
    
    Call PVSWcsv¼[ÌV[gì¬_Ver2001
    Call PVSWcsvÉTuio[ðnµÄTu}f[^ì¬_2017
    
    'gp·étB[h¼ÌZbg
    Dim fieldname As String: fieldname = "RLTFtoPVSW_,n_¤[¯Êq,I_¤[¯Êq,n_¤[îèiÔ,I_¤[îèiÔ,dã¡@_,Ú±G_,\¬_,¶æ_"
    ff = Split(fieldname, ",")
    Dim f As Variant: ReDim f(UBound(ff))
    For x = LBound(ff) To UBound(ff)
        f(x) = wb(0).Sheets("PVSW_RLTF").Cells.Find(ff(x), , , 1).Column
    Next x
    a = UBound(ff) + 1
    'düðZbg·ézñ
    Dim [düRAN As Variant
    ReDim [düRAN(a, 0)
    'tB[h¼ðzñÉüêé
    For x = LBound(ff) To UBound(ff)
        [düRAN(x, 0) = ff(x)
    Next x
    [düRAN(UBound([düRAN), 0) = "e[No"
    
    'ÎÛÌO[vÉ
    Dim CiÔi As Integer
    CiÔi = »iiÔRAN_read(»iiÔRan, "CiÔ")
    For i = LBound(»iiÔRan, 2) + 1 To UBound(»iiÔRan, 2)
        »iiÔstr = »iiÔRan(CiÔi, i)
        With wb(0).Sheets("PVSW_RLTF")
            '»iiÔÌtB[hðL[ÆµÄZbg
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(»iiÔstr, , , 1)
            Dim lastRow As Long
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For y = myKey.Row + 1 To lastRow
                If .Cells(y, myKey.Column) <> "" Then
                    If .Cells(y, f(0)) = "Found" Then
                        ReDim Preserve [düRAN(a, UBound([düRAN, 2) + 1)
                        For x = LBound(f) + 1 To UBound(f)
                            [düRAN(x, UBound([düRAN, 2)) = .Cells(y, f(x))
                        Next x
                        '[düRAN(0, UBound([düRAN, 2)) = 1
                    End If
                End If
            Next y
        End With
       Call ReplaceLR([düRAN)
    
        Call SumRan([düRAN) '¼[Ìs«æª¯¶AÚ±Gª¯¶êÜÆßé
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
                For x = 1 To 2
                    If [str = [düRAN(ii, x) Then
                        [düRAN(ii, UBound([düRAN, 2)) = Tustr
                        Exit For
                    End If
                Next x
            Next ii
        End If
    Next i
    
    'Call export_ArrayToSheet([düRAN, "[düRAN", False)
    'todo
    'Tuio[ðzz
    For ii = LBound([düRAN) To UBound([düRAN)
        'Ú±GÉæé»f
        Select Case Left([düRAN(ii, 6), 1)
            Case "T"
                '½àµÈ¢
            Case "E", "J", "B"
                [düRAN(ii, UBound([düRAN, 2)) = "999"
            Case "W"
                [düRAN(ii, UBound([düRAN, 2)) = "999"
        End Select
        '¶æ_Éæé»f
        Select Case Left([düRAN(ii, 8), 1)
            Case "E"
                [düRAN(ii, UBound([düRAN, 2)) = "999"
        End Select
    Next ii
    'Call export_ArrayToSheet([düRAN, "[düRAN", False)
    
    ']¿Ì¢eðîÉTuio[ðzzµÄ¢­
    For i = LBound([]¿RAN) + 1 To UBound([]¿RAN)
        [str = []¿RAN(i, 0)
        è[str = []¿RAN(i, 6)
        'If [str = "250" Then Stop
        If []¿RAN(i, 5) <> "" Then GoTo line20
        []¿RAN(i, 5) = [str
        For j = LBound([düRAN) + 1 To UBound([düRAN)
'            If j = 227 Then Stop
'            If i = 3 And j = 207 Then Stop
            If [düRAN(j, UBound([düRAN, 2)) = "" Then 'Ü¾Tuio[ªÜÁÄ³¯êÎ
                For x = 1 To 2
                    If [str = [düRAN(j, x) Then
                        []¿lng = [düRAN(j, 0)
                        If x = 1 Then è[str = [düRAN(j, 2)
                        If x = 2 Then è[str = [düRAN(j, 1)
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
                            If [düRAN(j, UBound([düRAN, 2)) = "" Then
                                [düRAN(j, UBound([düRAN, 2)) = [str
                            End If
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
                Next x
            End If
        Next j
line20:
    Next i
    
    'Call export_ArrayToSheet([]¿RAN, "[]¿RAN", False)
    
    'Call export_ArrayToSheet([düRAN, "[düRAN", False)
    'qªçÈ©Á½düð]¿Ì¢[É}·æ¤É·é
    For i = LBound([düRAN) + 1 To UBound([düRAN)
        If [düRAN(i, 0) <> "" Then
            If [düRAN(i, UBound([düRAN, 2)) = "" Then
                [1str = [düRAN(i, 1)
                [2str = [düRAN(i, 2)
                []¿1str = search[]¿RAN([]¿RAN, [1str, 2)
                []¿2str = search[]¿RAN([]¿RAN, [2str, 2)
                If []¿1str > []¿2str Then
                    ®Tustr = search[]¿RAN([]¿RAN, [1str, 5)
                Else
                    ®Tustr = search[]¿RAN([]¿RAN, [2str, 5)
                End If
                [düRAN(i, UBound([düRAN, 2)) = ®Tustr
            End If
        End If
    Next i
    
    'Call export_ArrayToSheet([düRAN, "[düRAN", False)
    
    '[888Ìzz
    Dim [str1 As String, [str2 As String, Ú±Gstr As String
    For i = LBound([düRAN) + 1 To UBound([düRAN)
        [str1 = [düRAN(i, 1)
        [str2 = [düRAN(i, 2)
        If [str1 & [str2 = "" Then [düRAN(i, UBound([düRAN, 2)) = "999"
    Next i
    
'    Call export_ArrayToSheet([düRAN, "[düRAN", False)
    
    'RLFTtoPVSW_ªóÌêO·é
    For i = LBound([düRAN) To UBound([düRAN)
        If i > UBound([düRAN) Then Exit For
        If [düRAN(i, 0) = "" Then
            [düRAN = removeArrayIndex([düRAN, i)
        End If
    Next i
    
   'Call export_ArrayToSheet([düRAN, "[düRAN", False)
    
    '[ÌTuio[ªdüÌTuio[É³¢êAcÉ·é
    Dim foundFlg As Boolean
    For i = LBound([]¿RAN) + 1 To UBound([]¿RAN)
        foundFlg = False
        Tustr = []¿RAN(i, 5)
        For ii = LBound([düRAN) + 1 To UBound([düRAN)
            If Tustr = [düRAN(ii, UBound([düRAN, 2)) Then
                foundFlg = True
                Exit For
            End If
        Next ii
        If foundFlg = False Then
            []¿RAN(i, 5) = "c"
        End If
    Next i
    
    '[]¿RANðeLXgoÍ
    Dim myTextPath As String
    myTextPath = wb(0).path & dirString_09
    makeDir myTextPath
    myTextPath = myTextPath & Replace(»iiÔstr, " ", "") & "_term.txt"
    export_Array_ShiftJis []¿RAN, myTextPath, ","
    
    '[düRANðeLXgoÍ
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(»iiÔstr, " ", "") & "_wiresum.txt"
    export_Array_ShiftJis [düRAN, myTextPath, ","
    
    'Call export_ArrayToSheet([düRAN, "[düRAN", False)
    
    'düÉTuio[(ìÆ)ÆXebvio[ðßé
    Dim myRan As Variant
    myRan = setWorkRanV2(»iiÔstr)
    
'    Call export_ArrayToSheet([]¿RAN, "[]¿RAN", False)
    
    'düÌTuio[ð[ÌTuio[Én·
    Dim subNumber As String, e[str As String
    For y = LBound([]¿RAN) To UBound([]¿RAN)
        e[str = []¿RAN(y, 5)
        If e[str = "c" Then
            subNumber = "c"
        Else
            subNumber = searchRan_ver2(myRan, e[str, "e[No", "subNumber")
        End If
        []¿RAN(y, UBound([]¿RAN, 2)) = subNumber
    Next y
    []¿RAN = WorksheetFunction.transpose([]¿RAN)
    
    'Call export_ArrayToSheet(myRan, "myRan", True)
    
    'PVSW_RLTFÉTuio[ðzz
    For i = LBound(»iiÔRan, 2) + 1 To UBound(»iiÔRan, 2)
        »iiÔstr = »iiÔRan(CiÔi, i)
        With wb(0).Sheets("PVSW_RLTF")
            '»iiÔÌtB[hðL[ÆµÄZbg
            Set myKey = wb(0).Sheets("PVSW_RLTF").Cells.Find(»iiÔstr, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For y = myKey.Row + 1 To lastRow
                If .Cells(y, myKey.Column) <> "" Then
                    If .Cells(y, f(0)) = "Found" Then
                        \¬str = Left(.Cells(y, f(7)), 4)
                        subNumber = searchRan_ver2(myRan, \¬str, "\¬_", "subNumber")
                        .Cells(y, myKey.Column) = subNumber
                        .Cells(y, myKey.Column).Interior.color = theme_color1
                    End If
                End If
            Next y
        End With
    Next i
    
    '[êÉTuio[ðzz
    For i = LBound(»iiÔRan, 2) + 1 To UBound(»iiÔRan, 2)
        »iiÔstr = »iiÔRan(CiÔi, i)
        Set ws(3) = wb(0).Sheets("[ê")
        With ws(3)
            Dim myCol(1) As Integer
            myCol(0) = .Cells.Find("[îèiÔ", , , 1).Column
            myCol(1) = .Cells.Find("[", , , 1).Column
            '»iiÔÌtB[hðL[ÆµÄZbg
            Set myKey = ws(3).Cells.Find(»iiÔstr, , , 1)
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            For y = myKey.Row + 1 To lastRow
                If .Cells(y, myKey.Column) <> "" Then
                    [îèiÔstr = .Cells(y, myCol(0))
                    [str = .Cells(y, myCol(1))
                    subNumber = searchRan_ver2([]¿RAN, [str & "," & [îèiÔstr, "[No,[îèiÔ", "subNumber")
                    If subNumber = "" Then Stop
                    .Cells(y, myKey.Column) = subNumber
                    .Cells(y, myKey.Column).Interior.color = theme_color1
                End If
            Next y
        End With
    Next i
    
    'zñðeLXgt@CoÍ·é
    '[]¿RANðeLXgoÍ
    []¿RAN = WorksheetFunction.transpose([]¿RAN)
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(»iiÔstr, " ", "") & "_term.txt"
    export_Array_ShiftJis []¿RAN, myTextPath, ","
    
    '[düRANðeLXgoÍ
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(»iiÔstr, " ", "") & "_wireSum.txt"
    export_Array_ShiftJis [düRAN, myTextPath, ","
    
    'myRANðeLXgoÍ
    myRan = WorksheetFunction.transpose(myRan)
    myTextPath = wb(0).path & "\09_AutoSub\"
    makeDir myTextPath
    myTextPath = myTextPath & Replace(»iiÔstr, " ", "") & "_wire.txt"
    export_Array_ShiftJis myRan, myTextPath, ","
    
    Call ÅK»àÇ·
    PlaySound "©ñ¹¢"
'
'    addRow = 1
'    For y = LBound([düRAN) To UBound([düRAN)
'            For x = LBound([düRAN, 2) To UBound([düRAN, 2)
'                With Sheets("temp")
'                    .Cells(addRow, x + 1) = [düRAN(y, x)
'                End With
'            Next x
'            addRow = addRow + 1
'    Next y
'
'    addRow = 1
'    For y = LBound([]¿RAN) To UBound([]¿RAN)
'        For x = LBound([]¿RAN, 2) To UBound([]¿RAN, 2)
'            With Sheets("temp2")
'                .Cells(addRow, x + 1) = []¿RAN(y, x)
'            End With
'        Next x
'        addRow = addRow + 1
'    Next y

    Dim myMsg As String: myMsg = "µÜµ½" & vbCrLf & DateDiff("s", mytime, time) & "s"
    aa = MsgBox(myMsg, vbOKOnly, "¶Yõ+Tu©®§Ä")
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
