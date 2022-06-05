Attribute VB_Name = "Module00"
Public SjpSetting_list As String

Public Function BIN2HEX(myBIN)
    If Len(myBIN) Mod 4 > 0 Then
        myBIN = String(((Len(myBIN) \ 4) + 1) * 4 - Len(myBIN), "0") & myBIN
    End If
    
    For u = 1 To Len(myBIN) Step 4
        Select Case Mid(myBIN, u, 4)
            Case "0000"
            myHEX = myHEX & "0"
            Case "0001"
            myHEX = myHEX & "1"
            Case "0010"
            myHEX = myHEX & "2"
            Case "0011"
            myHEX = myHEX & "3"
            Case "0100"
            myHEX = myHEX & "4"
            Case "0101"
            myHEX = myHEX & "5"
            Case "0110"
            myHEX = myHEX & "6"
            Case "0111"
            myHEX = myHEX & "7"
            Case "1000"
            myHEX = myHEX & "8"
            Case "1001"
            myHEX = myHEX & "9"
            Case "1010"
            myHEX = myHEX & "A"
            Case "1011"
            myHEX = myHEX & "B"
            Case "1100"
            myHEX = myHEX & "C"
            Case "1110"
            myHEX = myHEX & "D"
            Case "1111"
            myHEX = myHEX & "F"
        End Select
    Next u
    BIN2HEX = myHEX
End Function

Public Function 原紙の設定(myBook, 原紙, 保存フォルダ名, newBookName) As Workbook

    拡張子 = Mid(原紙, InStrRev(原紙, "."))
    newBookName = Left(myBook.Name, InStrRev(myBook.Name, ".") - 1) & "_" & newBookName
    
    '重複しないファイル名に決める
    For i = 0 To 999
        If Dir(wb(0).path & "\" & 保存フォルダ名 & "\" & newBookName & "_" & Format(i, "000") & 拡張子) = "" Then
            newBookName = newBookName & "_" & Format(i, "000")
            Exit For
        End If
        If i = 999 Then Stop '想定していない数
    Next i
    
    '原紙を読み取り専用で開く
    On Error Resume Next
    Workbooks.Open fileName:=myAddress(0, 1) & "\" & 原紙, ReadOnly:=True
    On Error GoTo 0
    
    '原紙をサブ図のファイル名に変更して保存
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=wb(0).path & "\" & 保存フォルダ名 & "\" & newBookName & 拡張子
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set 原紙の設定 = ActiveWorkbook
End Function

Public Function オートシェイプ削除()
    'オートシェイプを削除
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        objShp.Delete
    Next objShp
End Function

Public Function 製品使分け結合(結合前1, 結合前2)
    結合前1s = Split(結合前1, ",")
    結合前2s = Split(結合前2, ",")
    
    For i = LBound(結合前1s) To UBound(結合前1s)
        If 結合前1s(i) <> "" Then
            結合後 = 結合後 & "," & 結合前1s(i)
        Else
            結合後 = 結合後 & "," & 結合前2s(i)
        End If
    Next i
    
    製品使分け結合 = right(結合後, Len(結合後) - 1)
    
End Function

Public Function 冶具図_端末経路表示()
    Call 最適化
    Set myBook = ActiveWorkbook
    Dim 端末str As String
    端末str = Application.Caller
    On Error Resume Next
    ActiveSheet.Shapes("配索").Ungroup
    ActiveSheet.Shapes("冶具").Ungroup
    On Error GoTo 0
    Call SQL_配索端末取得_端末用端末(配索端末RAN, 端末str)

    For Each ob In ActiveSheet.Shapes
        If InStr(ob.Name, "!") Then
            ob.Delete
        Else
            If ob.Type = 1 Then
                ob.Line.ForeColor.RGB = RGB(0, 0, 0)
                ob.Fill.ForeColor.RGB = RGB(255, 255, 255)
            ElseIf ob.Type = 9 Then
                ob.Line.ForeColor.RGB = RGB(150, 150, 150)
            End If
        End If
    Next ob
    Dim 配索toStr As String
    With ActiveSheet
        '■選択した端末の色付け
        With .Shapes(端末str)
            .Select
            .ZOrder msoBringToFront
            .Fill.ForeColor.RGB = RGB(255, 100, 100)
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame.Characters.Font.color = RGB(0, 0, 0)
            '.Line.Weight = 2
            myTop = Selection.Top
            myLeft = Selection.Left
            myHeight = Selection.Height
            myWidth = Selection.Width
        End With

        '■配索する端末間のラインに色付け
        Set 端末from = .Cells.Find(端末str, , , 1)
        For i = LBound(配索端末RAN) To UBound(配索端末RAN)
            Dim myStep As Long
            端末to = 配索端末RAN(i)
            If 端末to = "" Then GoTo nextI
            Set 配索 = .Cells.Find(端末str, , , 1)
            If 配索 Is Nothing Then GoTo nextI
                Set 端末to = .Cells.Find(配索端末RAN(i), , , 1)
                If 端末to Is Nothing Then GoTo nextI
                If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                ActiveSheet.Shapes(端末to.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                ActiveSheet.Shapes(端末to.Value).ZOrder msoBringToFront
                配索toStr = 配索toStr & "," & 端末to.Value
                Set 端末1 = 端末from
                For y = 端末from.Row To 端末to.Row Step myStep
                    'fromから左に進む
                    Do Until 端末1.Column = 1
                        Set 端末2 = 端末1.Offset(0, -2)
                        On Error Resume Next
                            ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).ZOrder msoBringToFront
                            ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).ZOrder msoBringToFront
                        On Error GoTo 0
                        Set 端末1 = 端末2
                        If Left(端末1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                    Loop
                    
line15:
                    'toの行まで上または下に進む
                    Do Until 端末1.Row = 端末to.Row
                        Set 端末2 = 端末1.Offset(myStep, 0)
                        If 端末1 <> 端末2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).ZOrder msoBringToFront
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).ZOrder msoBringToFront
                            On Error GoTo 0
                        End If
                        Set 端末1 = 端末2
                        If Left(端末1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                    Loop
                    
                    'toの行を右に進む
                    Do Until 端末1.Column = 端末to.Column
                        Set 端末2 = 端末1.Offset(0, 2)
                        On Error Resume Next
                            ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Line.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).ZOrder msoBringToFront
                            ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).ZOrder msoBringToFront
                        On Error GoTo 0
                        Set 端末1 = 端末2
                        If Left(端末1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                    Loop
                Next y
                Set 端末2 = Nothing
nextI:
        Next i

        For Each ob In ActiveSheet.Shapes
            If ob.Type = 1 And ob.Name <> "板" Then
                ob.ZOrder msoBringToFront
            Else
                
            End If
        Next ob
        '■配索する後ハメ電線を表示
        Dim 配索toStrSp
        配索toStrSp = Split(配索toStr, ",")
        Dim 色v As String, サv As String, 端末v As String, マv As String, ハメv As String
        For ii = LBound(配索toStrSp) + 1 To UBound(配索toStrSp)
            端末v = 配索toStrSp(ii) '端末v=行き先
            Call SQL_配索端末取得_端末用回路(配索端末RAN, 端末v, 端末str)
            For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
                色v = 配索端末RAN(2, i)
                If 色v = "" Then Exit For
                マv = 配索端末RAN(6, i)
                サv = 配索端末RAN(4, i)
                ハメv = 配索端末RAN(4, i)
                構成v = 配索端末RAN(3, i)
                名前c = 0
                For Each objShp In ActiveSheet.Shapes
                    If objShp.Name = 端末v & "!" Then
                        名前c = 名前c + 1
                    End If
                Next objShp
                    
                With .Shapes(端末v)
                    .Select
                    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Selection.Left + Selection.Width + (名前c * 15), Selection.Top, 15, 15).Select
                    Call 色変換(色v, clocode1, clocode2, clofont)
                    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Left(Replace(サv, "F", ""), 3)
                    Selection.ShapeRange.Adjustments.Item(1) = 0.15
                    'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
                    Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.4
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.401
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.599
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.6
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.99
                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                    Selection.ShapeRange.Name = 端末v & "!"
                    If InStr(色v, "/") > 0 Then
                        ベース色 = Left(色v, InStr(色v, "/") - 1)
                    Else
                        ベース色 = 色v
                    End If
                    
                    myFontColor = clofont 'フォント色をベース色で決める
                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
                    Selection.ShapeRange.TextFrame2.TextRange.Font.size = 6
                    Selection.Font.Name = myFont
                    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                    Selection.ShapeRange.TextFrame2.MarginLeft = 0
                    Selection.ShapeRange.TextFrame2.MarginRight = 0
                    Selection.ShapeRange.TextFrame2.MarginTop = 0
                    Selection.ShapeRange.TextFrame2.MarginBottom = 0
                    'ストライプは光彩を使う
                    If clocode1 <> clocode2 Then
                        With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                            .color = clocode1
                            .color.TintAndShade = 0
                            .color.Brightness = 0
                            .Transparency = 0#
                            .Radius = 8
                        End With
                    End If
                    'マルマ
                    If マv <> "" Then
                        myLeft = Selection.Left
                        myTop = Selection.Top
                        myHeight = Selection.Height
                        myWidth = Selection.Width
                        For Each objShp In Selection.ShapeRange
                            Set objShpTemp = objShp
                        Next objShp
                        ActiveSheet.Shapes.AddShape(msoShapeOval, myLeft + (myWidth * 0.6), myTop + (myHeight * 0.6), myWidth * 0.4, myHeight * 0.4).Select
                        Call 色変換(マv, clocode1, clocode2, clofont)
                        myFontColor = clofont
                        Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
                        Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                        objShpTemp.Select False
                        Selection.Group.Select
                        Selection.Name = 端末v & "!"
                    End If
                End With
            Next i
        Next ii
    End With
    Call 最適化もどす
End Function

Public Function 配列を入れ替える(データ)
    '製品品番毎の製品使分けに置き換える_サブ№を1に置き換える
    Dim 配列() As String
    ReDim 配列(1, 製品品番RANc - 1) '0:電線使分け,1:製品使分け
    
    For i = LBound(データ, 3) To UBound(データ, 3)
        データs = Split(データ(1, 1, i), ",")
        For a = LBound(データs) To UBound(データs)
            If データs(a) <> "" Then
                配列(0, a) = 配列(0, a) & ",1"
            Else
                配列(0, a) = 配列(0, a) & ",0"
            End If
        Next a
    Next i
    '余分な","を削除
    For i = LBound(配列, 2) To UBound(配列, 2)
        配列(0, i) = right(配列(0, i), Len(配列(0, i)) - 1)
    Next i
    '電線があれば製品品番をセットする
    For i = LBound(配列, 2) To UBound(配列, 2)
        If InStr(配列(0, i), "1") > 0 Then 配列(1, i) = 製品品番Ran(1, i)
    Next i
    '電線使分けが同じ時は、片方を削除する
    For i = LBound(配列, 2) To UBound(配列, 2)
        If 配列(0, i) <> "0" Then
            For i2 = i To UBound(配列, 2)
                If i <> i2 Then
                        If 配列(0, i) = 配列(0, i2) Then
                            配列(0, i2) = ""
                            配列(1, i) = 配列(1, i) & "," & 配列(1, i2)
                            配列(1, i2) = ""
                        End If
                End If
            Next i2
        End If
    Next i
    配列を入れ替える = 配列
End Function

Public Function ソート0(newSheet, startRow, lastRow, 優先1, 優先2, 優先3)
    'ソート
    With newSheet
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, 優先1).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先2).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先3).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange .Range(.Rows(startRow), .Rows(lastRow))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
End Function

Sub DeleteDefinedNames()
 
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        If n.MacroType = -4142 Then
            n.Delete
        End If
    Next
 
End Sub
Public Function 製品品番RAN_read(製品品番Ran, 製品品番FIE)

    For i = LBound(製品品番Ran, 1) To UBound(製品品番Ran, 1)
        If 製品品番Ran(i, 0) = 製品品番FIE Then
            製品品番RAN_read = i
            Exit Function
        End If
    Next i

End Function
Public Function 製品品番RAN_seek()
    For x = 1 To 製品品番Rc
        If 製品品番Ran(1, x) = "" Then Stop '製品品番がセットされてないと探せない
        For xx = 1 To 製品品番RANc
            If 製品品番Ran(1, x) = 製品品番Ran(1, xx) Then
                For a = 1 To 10
                    製品品番Ran(a, x) = 製品品番Ran(a, xx)
                Next a
                GoTo line10
            End If
        Next xx
        Stop '製品品番が見つからない
line10:
    Next x
End Function
Public Function ProgressBar_ref(処理名 As String, 処理内容 As String, step0T As Long, step0 As Long, Step1T As Long, Step1 As Long)
    With ProgressBar
        .Caption = "処理中 " & 処理名
        
        .ProgBar0.Max = step0T
        .ProgBar0.Value = step0
        .msg0.Caption = step0 & "/" & step0T & "  " & 処理内容
        
        .ProgBar1.Max = Step1T
        .ProgBar1.Value = Step1
        .msg1.Caption = Step1 & "/" & Step1T
        '.Repaint
        DoEvents
        'If .StopBtn.Value = True Then Stop
        
    End With
End Function
Public Function コメント表示切替()
    Dim コメント表示 As Boolean
    With Sheets("設定")
        コメント表示 = .Cells.Find("コメント表示切替", , , 1).Offset(0, 1).Value
    End With
    
    コメント表示 = コメント表示 + 1
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
        
    For Each cmt In ws.Comments
        cmt.Visible = コメント表示
    Next cmt
    
    With Sheets("設定")
        .Cells.Find("コメント表示切替", , , 1).Offset(0, 1) = コメント表示
    End With
End Function

Public Function 完了しました(Optional myBook)
    myBook.Activate
    'Set aa = ActiveSheet.Shapes.AddPicture("H:\作成しました.png", False, True, 0, 0, 164, 128)
    Set aa = ActiveSheet.Pictures.Insert(myAddress(0, 1) & "\picture\作成しました.png")
    winW = Application.Width
    winH = Application.Height
    aa.Left = (winW - aa.Width) / 2
    aa.Top = (winH - aa.Height) / 2
    aa.OnAction = "閉じる"
    PlaySound ("かんせい")
End Function

Public Function 閉じる()
    Set myBook = ActiveWorkbook
    myme = Application.Caller
    ActiveSheet.Shapes(myme).Delete
    PlaySound ("とじる2")
    ActiveWorkbook.VBProject.VBComponents(ActiveSheet.codeName).CodeModule.AddFromFile myAddress(0, 1) & "\OnKey\002_問連書作成_マルマ.txt"
    Application.OnKey "^{ENTER}", "問題点連絡書_マルマ_Ver2002"
    Application.OnKey "^~", "問題点連絡書_マルマ_Ver2002"
End Function

Public Sub addressSet(ByVal myBook As Workbook)
    ReDim myAddress(4, 1)
    Set wb(0) = myBook
    If myAddress(0, 0) = "" And myBook Is Nothing Then Set wb(0) = ActiveWorkbook
    
    If myIP = "" Then Exit Sub
    Dim myIPs As String, addressWords As String, sp As Variant, i As Long, tempPath As String
    myIPs = Mid(myIP, 1, InStrRev(myIP, ".") - 1)
    addressWords = "システムパーツ_,部材一覧+_," & myIPs & "_wireSub.txt," & myIPs & "_tubeSub.txt," & myIPs & "_partsSub.txt"
    sp = Split(addressWords, ",")
    For i = LBound(sp) To UBound(sp)
        tempPath = sp(i)
        myAddress = addPath(myAddress, myBook.Sheets("設定"), tempPath, i)
    Next i
    
End Sub

Public Function addPath(ByVal myRan As Variant, ByVal ws As Worksheet, ByVal searchWord As String, ByVal i As Long) As Variant
    Dim firstCell As Variant, foundCell As Variant, tempCount As Long, y As Long
    
    With ws
        '数を調べる
        tempCount = 0
        SjpSetting_list = ""
        Set firstCell = .Cells.Find(searchWord, , , 1)
        Set foundCell = firstCell
        Do
            Set foundCell = .Cells.FindNext(foundCell)
            tempCount = tempCount + 1
            SjpSetting_list = SjpSetting_list & "," & foundCell.Offset(0, 1).Value
        Loop While firstCell.Row <> foundCell.Row
        
        If Not foundCell Is Nothing Then
            If tempCount = 1 Then
                Set firstCell = .Cells.Find(searchWord, , , 1)
                Set foundCell = firstCell
                myRan(i, 0) = searchWord
                myRan(i, 1) = foundCell.Offset(0, 1)
            Else
                '複数ある場合
                Dim sp As Variant
                SjpSetting_list = Mid(SjpSetting_list, 2)
                sp = Split(SjpSetting_list, ",")
                
                Dim SjpSettingArray As Variant
                SjpSettingArray = readTextToArray(SjpSetting_Path)
                
                'ファイルが無い場合
                If Dir(SjpSetting_Path) = "" Then UI_12.Show
                
                Dim foundFlag As Boolean: foundFlag = False
                For y = LBound(sp) To UBound(sp)
                    If sp(y) = SjpSettingArray(0, 1) Then
                        myRan(i, 0) = searchWord
                        myRan(i, 1) = SjpSettingArray(0, 1)
                        foundFlag = True
                    End If
                Next y
                '値が見つからない場合
                If foundFlag = False Then UI_12.Show
            End If
        End If
    End With
    
    addPath = myRan
End Function

Public Function 参照不可があればそのフォルダを作成する()

    Dim Ref, buf As String, bufS, myCount As Long
    Dim myProject(8) As String
    myProject(0) = ""            'VBEのバージョンによるので使用しない_VBE7.DLL
    myProject(1) = ""            'EXCEL.EXEのバージョンによるので使用しない_Office15
    myProject(2) = "stdole2.tlb"
    myProject(3) = "MSO.DLL"
    myProject(4) = "scrrun.dll"
    myProject(5) = "FM20.DLL"
    myProject(6) = "msado15.dll"
    myProject(7) = "REFEDIT.DLL"
    myProject(8) = "MSCOMCTL.OCX"
    
    '参照不可がある場合bufにセットする
    For Each Ref In wb(0).VBProject.References
        If Ref.isbroken = True Then
            buf = buf & myCount & vbTab & Ref.Name & vbTab & Ref.Description & vbTab & Ref.FullPath & vbCrLf
        End If
        myCount = myCount + 1
    Next Ref
    
    Debug.Print buf
    '参照不可がある場合
    If buf <> "" Then
        bufS = Split(buf, vbCrLf)
        For i = LBound(bufS) To UBound(bufS) - 1
            bufss = Split(bufS(i), vbTab)
            'フォルダが無ければ作成
            dirsp = Split(bufss(3), "\")
            dirstr = ""
            For i2 = LBound(dirsp) To UBound(dirsp) - 1
                dirstr = dirstr & "\" & dirsp(i2)
                If Dir(Mid(dirstr, 2), vbDirectory) = "" Then
                    MkDir Mid(dirstr, 2)
                End If
            Next i2
        Next i
        'ライブラリファイルのコピー
        If Dir(bufss(3)) = "" Then
            FileCopy myAddress(0, 1) & "\DLL\" & myProject(bufss(0)), bufss(3)
        End If
    End If
End Function
'参照設定が無ければ追加、多い場合の処理は書いてない
Sub 参照設定の変更()
    '参照設定一覧を取得する時用
'    Dim Ref2, buf2 As String
'    For Each Ref2 In ActiveWorkbook.VBProject.References
'        buf2 = buf2 & Ref2.Name & vbLf & Ref2.Description & vbLf & Ref2.FullPath & vbLf & vbLf
'    Next Ref2
'    MsgBox buf2
'    Stop
    
    Set wb(0) = ThisWorkbook
    '対象の参照設定の一覧。wb(0).VBProject.References.Description,Fullpath
    Dim Ref(1, 2) As String, myFlg As Boolean
    Ref(0, 0) = "Microsoft Windows Common Controls 6.0 (SP6)": Ref(1, 0) = "C:\Windows\SysWow64\MSCOMCTL.OCX"
    Ref(0, 1) = "Ref Edit Control": Ref(1, 1) = "C:\Program Files (x86)\Microsoft Office\Office14\REFEDIT.DLL"
    Ref(0, 2) = "Windows Script Host Object Model": Ref(1, 2) = "C:\Windows\SysWOW64\wshom.ocx"
    For i = LBound(Ref, 2) To UBound(Ref, 2)
        '現在の参照設定をチェック
        myFlg = False
        For Each buf In wb(0).VBProject.References
            If buf.Description = Ref(0, i) Then myFlg = True
        Next buf
        If myFlg = False Then wb(0).VBProject.References.AddFromFile Ref(1, i)
    Next i
    
End Sub

Public Function ハメ色変更()
    Dim keyRow As Long, keyCol As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim 側 As String
    Dim フィールド名(2) As String
    フィールド名(0) = "側回路符号"
    フィールド名(1) = "側端末識別子"
    フィールド名(2) = "側キャビティ"
    
    Dim 色 As Variant: 色 = RGB(0, 102, 0)
    
    With ActiveSheet
        keyRow = .Cells.Find("電線識別名", , , 1).Row
        x = ActiveCell.Column
        'y = ActiveCell.Row
        側 = Left(.Cells(keyRow, x).Value, 2)
        If 側 = "始点" Or 側 = "終点" Then
            For y = Selection(1).Row To Selection(Selection.count).Row
                If y <= keyRow Then GoTo line10
                For i = LBound(フィールド名) To UBound(フィールド名)
                    keyCol = .Cells.Find(側 & フィールド名(i), , , 1).Column
                    .Cells(y, keyCol).Font.color = 色
                    .Cells(y, keyCol).Font.Bold = True
                Next i
line10:
            Next y
        Else
            Exit Function
        End If
    End With

End Function

Public Function ディレクトリ作成()
    Dim myDir As String, myDirS As Variant
    
    myDir = "\01_PVSW_csv,\05_RLTF_A,\06_RLTF_B,\07_SUB,\08_MD,\08_MD,\08_hsfデータ変換,\08_hsfデータ変換\log,\A0_部材一覧"
    
    myDirS = Split(myDir, ",")
    For i = LBound(myDirS) To UBound(myDirS)
        If Dir(ActiveWorkbook.path & "\" & myDirS(i), vbDirectory) = "" Then
            MkDir ActiveWorkbook.path & myDirS(i)
        End If
    Next i
End Function
Public Function 必要ファイルの取得()
    'exe
    Dim myDir As String, myDirS As Variant
    myDir = "\08_hsfデータ変換\WH_DataConvert.exe"
    myDirS = Split(myDir, ",")
    For i = LBound(myDirS) To UBound(myDirS)
        If Dir(ActiveWorkbook.path & "\" & myDirS(i)) = "" Then
            FileCopy myAddress(0, 1) & "\hsfデータ変換\WH_DataConvert.exe", ActiveWorkbook.path & "\" & myDirS(i)
        End If
    Next i
    'iniファイルを毎回作成しなおす
    Open ActiveWorkbook.path & "\08_hsfデータ変換\HsfDataConvert.ini" For Output As #1
        Print #1, "[Data]"
        Print #1, "HsfDataPath=" & ActiveWorkbook.path & "\08_hsfデータ変換"
        Print #1, "GuideDataPath=" & ActiveWorkbook.path & "\08_MD"
        Print #1, "HsfSearchCnt=200"
        Print #1, "HsfReadState=0"
        Print #1, "[Time]"
        Print #1, "StartHour=0"
        Print #1, "StartMin=0"
        Print #1, "StartSec=0"
        Print #1, "EndHour=23"
        Print #1, "EndMin=0"
        Print #1, "EndSec=0"
    Close #1
    '部材一覧+があるかチェック
    Dim buf As String, cnt As Long
    Dim Path1 As String: Path1 = ActiveWorkbook.path & "\" & "A0_部材一覧\Bip"
    buf = Dir(Path1 & "*.xlsm")
    Do While buf <> ""
        cnt = cnt + 1
        buf = Dir()
    Loop
    '部材一覧+が無い場合は最新版を取得
    If cnt = 0 Then
        Dim Path2 As String: Path2 = myAddress(1, 1) & "\down\Bip\"
        buf = Dir(Path2 & "*.xlsm")
        Dim thisVer As String, newVer As String, fileName As String
        Do While buf <> ""
            If Left(buf, 3) = "Bip" Then
                thisVer = Mid(buf, 4, InStr(buf, "_") - 4)
                If newVer = "" Then
                    newVer = thisVer
                Else
                    If thisVer > newVer Then
                        newVer = thisVer
                    End If
                End If
            End If
            buf = Dir()
        Loop
        FileCopy Path2 & "Bip" & newVer & "_.xlsm", Path1 & "Bip" & newVer & "_.xlsm"
    End If
End Function
Public Sub 部材詳細_端子ファミリー(strFilePath, 端子ファミリー)
    Dim intCount As Integer
    Dim intNo As Integer
    Dim strFileName As String
    Dim strBuff As String, getFlg As Boolean
    
    ' ファイルオープン
    intNo = FreeFile()                      ' フリーファイルNoを取得
    Open strFilePath For Input As #intNo    ' ファイルをオープン
    
    ' ファイルの読み込み
    intCount = 0
    Do Until EOF(intNo)                     ' ファイルの最後までループ
        getFlg = False
        Line Input #intNo, strBuff          ' ファイルから一行読み込み
        For k = LBound(端子ファミリー, 2) To UBound(端子ファミリー, 2)
            If InStr(strBuff, "," & 端子ファミリー(0, k)) > 0 Then
                getFlg = True
                Exit For
            End If
        Next k
        
        If intCount = 0 Or getFlg = True Then
            ReDim Preserve strArray(intCount)   ' 配列長を変更
            strArray(intCount) = strBuff        ' 配列の最終要素に読み込んだ値を代入
            intCount = intCount + 1             ' 配列の要素数を加算
        End If
    Loop
    
    ' ファイルクローズ
    Close #intNo
    
    ' 読み込んだ値を確認
'    Dim i As Integer
'    For i = 0 To UBound(strArray)
'        Debug.Print strArray(i)
'    Next i
    
End Sub

Public Sub SUBデータ取得(SUBデータRAN, strFilePath)
    Dim intCount As Integer
    Dim intNo As Integer
    Dim strFileName As String
    Dim strBuff As String, getFlg As Boolean
    
    ' ファイルオープン
    intNo = FreeFile()                      ' フリーファイルNoを取得
    Open strFilePath For Input As #intNo    ' ファイルをオープン
    ReDim SUBデータRAN(0)
    ' ファイルの読み込み
    intCount = 0
    Do Until EOF(intNo)                     ' ファイルの最後までループ
        getFlg = False
        Line Input #intNo, strBuff          ' ファイルから一行読み込み
        ReDim Preserve SUBデータRAN(UBound(SUBデータRAN) + 1)
        SUBデータRAN(UBound(SUBデータRAN)) = strBuff
    Loop
    
    ' ファイルクローズ
    Close #intNo
    
    ' 読み込んだ値を確認
'    Dim i As Integer
'    For i = 0 To UBound(strArray)
'        Debug.Print strArray(i)
'    Next i
    
End Sub


Public Sub 端子ファミリー検索(myCell, 端子ファミリー)
    For i = LBound(strArray) To UBound(strArray)
        strArrayS = Split(strArray(i), ",")
        '部品品番のマッチ確認
        If myCell = Replace(strArrayS(0), "-", "") Then
            'ファミリー番号のマッチ確認
            For ii = LBound(端子ファミリー, 2) To UBound(端子ファミリー, 2)
                If Left(strArrayS(13), 3) = 端子ファミリー(0, ii) Then
                    If strArrayS(14) = 端子ファミリー(1, ii) Or "*" = 端子ファミリー(1, ii) Then
                        myCell.Interior.color = 端子ファミリー(3, ii)
                        '特記tempに登録があるか確認
                        Set fnd = Range("端子ファミリー範囲").Find(端子ファミリー(0, ii) & 端子ファミリー(1, ii), , , 1)
                        If fnd Is Nothing Then
                            For Each f In Range("端子ファミリー範囲")
                                If f.Value = "" Then
                                    Sheets("設定").Hyperlinks.add anchor:=f, addRess:=端子ファミリー(2, ii), ScreenTip:="", TextToDisplay:=端子ファミリー(0, ii) & 端子ファミリー(1, ii)
                                    f.Interior.color = 端子ファミリー(3, ii)
                                    f.Font.color = 0
                                    f.Font.Underline = False
                                    f.AddComment
                                    f.Comment.Text 端子ファミリー(5, ii)
                                    f.Comment.Shape.TextFrame.AutoSize = True
                                    Exit Sub
                                End If
                            Next f
                        End If
                    End If
                End If
            Next ii
            Exit Sub
        End If
    Next i
    '見つからなかった
    'Stop  '部材一覧の処理が未だ?
End Sub

Public Sub 電線品種検索(myCell, 電線品種)
    '電線品種のマッチ確認
    For ii = LBound(電線品種, 2) To UBound(電線品種, 2)
        If myCell = 電線品種(1, ii) Then
                myCell.Interior.color = 電線品種(3, ii)
                '電線品種tempに登録があるか確認
                Set fnd = Range("電線品種範囲").Find(電線品種(0, ii), , , 1)
                If fnd Is Nothing Then
                    For Each f In Range("電線品種範囲")
                        If f.Value = "" Then
                            Sheets("設定").Hyperlinks.add anchor:=f, addRess:=電線品種(2, ii), ScreenTip:="", TextToDisplay:=電線品種(0, ii)
                            f.Interior.color = 電線品種(3, ii)
                            f.Font.color = 0
                            f.Font.Underline = False
                            If 電線品種(5, ii) <> "" Then
                                f.AddComment
                                f.Comment.Text 電線品種(5, ii)
                                f.Comment.Shape.TextFrame.AutoSize = True
                            End If
                            Exit Sub
                        End If
                    Next f
                End If
        End If
    Next ii
End Sub

Public Function 部材詳細_set(strFilePath, filterWord, u, myX)
    Dim intCount As Integer
    Dim intNo As Integer
    Dim strFileName As String
    Dim strBuff As String, getFlg As Boolean
    
    ' ファイルオープン
    intNo = FreeFile()                      ' フリーファイルNoを取得
    Open strFilePath For Input As #intNo    ' ファイルをオープン
    
    ' ファイルの読み込み
    intCount = 0
    Do Until EOF(intNo)                     ' ファイルの最後までループ
        getFlg = False
        Line Input #intNo, strBuff          ' ファイルから一行読み込み
        'フィールド名を指定
        If intCount = 0 Then
            strbuffsp = Split(strBuff, ",")
            For i = LBound(strbuffsp) To UBound(strbuffsp)
                If strbuffsp(i) = filterWord Then
                    myX = i
                    Exit For
                End If
            Next i
        End If
        '登録する条件
        
        strbuffsp = Split(strBuff, ",")
        If strbuffsp(myX) <> "" Then
            '登録
            ReDim Preserve strArray(intCount)   ' 配列長を変更
            strArray(intCount) = strBuff        ' 配列の最終要素に読み込んだ値を代入
            intCount = intCount + 1             ' 配列の要素数を加算
        End If
    Loop
    
    ' ファイルクローズ
    Close #intNo

End Function

Public Function TEXT出力_汎用検査履歴システム(myDir, 構成, 色呼, サブ, point, 端末, 作業工程)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 14) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir & "\" & Format(point, "0000") & ".html"

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift_JIS" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=8" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./img/wh.css" & Chr(34) & ">"
        outdats(6) = "<title>" & 構成 & "</title>"
        outdats(7) = "</head>"
        outdats(8) = "<body>"
        outdats(9) = "<table>"
        outdats(10) = "<tr><td class=" & Chr(34) & "title" & Chr(34) & "> 構成:" & 構成 & " " & 色呼 & " 工程:" & サブ & " " & 作業工程 & "</td></tr>"
        outdats(11) = "<tr><td><img src=" & Chr(34) & "./img/" & Format(point, "0000") & ".jpg" & Chr(34) & "></td></tr>"
        outdats(12) = "</table>"
        outdats(13) = "</body>"
        outdats(14) = "</html>"
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber
    
End Function
Public Function TEXT出力_汎用検査履歴システムhtml(myDir, 構成, 色呼, サブ, point, 端末, 作業工程, cav)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 17) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir & "\" & Format(point, "0000") & ".html"

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift_JIS" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=8" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/wh" & Format(point, "0000") & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & point & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "myBlink()" & Chr(34) & " >"
        
        outdats(9) = "<table>"
        
        outdats(10) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>" & 構成 & "</td><td>" & 色呼 & "</td>"
        If 後ハメ作業者 = True Then outdats(10) = outdats(10) & "<td>" & myVer & " " & 後ハメ作業者シート名 & "</td>"
        outdats(10) = outdats(10) & "<td>" & サブ & "</td><td>" & 作業工程 & "</td></tr>"
        outdats(11) = "</table>"
                
        outdats(12) = "<div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末 & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " ></div>"
        outdats(13) = "<div id=" & Chr(34) & "box2" & Chr(34) & " ><img src=" & Chr(34) & "./img/" & 端末 & "_1_" & cav & ".png" & Chr(34) & "></div>"
        outdats(14) = ""
        
        outdats(15) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(16) = "</body>"
        outdats(17) = "</html>"
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber
    
    TEXT出力_汎用検査履歴システムhtml = myPath

End Function

Public Function TEXT出力_配索経路html(myDir, 端末from, 端末to, 製品品番str, サブ, サブ2, 構成, 色呼, 始点ハメ, 始点cav, 終点ハメ, 終点cav, 端末leftRAN)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 38) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir
    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift-jis" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/" & 構成 & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & 構成 & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "myBlink();myBlink2();document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>構成:" & 構成 & " " & 色呼 & "</td><td>" & 端末from & " to " & 端末to & "</td><td>SUB:" & サブ & "</td><td>Ver:" & myVer & "</td>" & _
                               "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                               "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
        '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        Dim 端末fromleft As Single, 端末toleft As Single, 端末from1 As String, 端末from2 As String, 端末to1 As String, 端末to2 As String
        端末fromleft = 0: 端末toleft = 0
        For i = LBound(端末leftRAN, 2) + 1 To UBound(端末leftRAN, 2)
            If 端末from = 端末leftRAN(0, i) Then 端末fromleft = 端末leftRAN(1, i)
            If 端末to = 端末leftRAN(0, i) Then 端末toleft = 端末leftRAN(1, i)
        Next i
        '右にある方を右に表示させるbox6と7だと右になる
        If val(端末fromleft) >= val(端末toleft) Then
            端末from1 = "box6"
            端末from2 = "box7"
            端末to1 = "box4"
            端末to2 = "box5"
        Else
            端末from1 = "box4"
            端末from2 = "box5"
            端末to1 = "box6"
            端末to2 = "box7"
        End If
        
        If Left(始点ハメ, 1) = "後" Then
            outdats(14) = "  <div  id=" & Chr(34) & 端末from1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_1.png" & Chr(34) & " ></div>"
            outdats(15) = "  <div  id=" & Chr(34) & 端末from2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_1_" & 始点cav & ".png" & Chr(34) & " ></div>"
        End If
        
        If Left(終点ハメ, 1) = "後" Then
            outdats(16) = "  <div id=" & Chr(34) & 端末to1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末to & "_1.png" & Chr(34) & " ></div>"
            outdats(17) = "  <div id=" & Chr(34) & 端末to2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末to & "_1_" & 終点cav & ".png" & Chr(34) & " ></div>"
        End If
        outdats(18) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & サブ2 & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"
        outdats(19) = "  <div id=" & Chr(34) & "box2" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "to" & 端末to & "_" & 色呼 & ".png" & Chr(34) & " ></div>"
        outdats(20) = "  <div id=" & Chr(34) & "box3" & Chr(34) & "><img src=" & Chr(34) & "./img/" & サブ & "_foot.png" & Chr(34) & " ></div>"
        outdats(21) = "</body>"
        
        outdats(22) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(23) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink2.js" & Chr(34) & "></script>"
        outdats(24) = "<script>"
        outdats(25) = "function checkText(){"
        outdats(26) = "  var str1=document.myform.txtb.value;"
        outdats(27) = "  var seihin,kosei;"
        outdats(28) = "  var myLen=str1.length;"
        outdats(29) = "  if (myLen <=10){"
        outdats(30) = "    kosei=str1;"
        outdats(31) = "  }else{"
        outdats(32) = "    seihin=str1.substr(25,10);"
        outdats(33) = "    kosei=str1.substr(11,4);"
        outdats(34) = "  }"
        outdats(35) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(36) = "}"
        outdats(37) = "</script>"
        outdats(38) = "</html>"
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function

Public Function TEXT出力_配索経路html_UTF8(myDir, 端末from, 端末to, 製品品番str, サブ, サブ2, 構成, 色呼, 始点ハメ, 始点cav, 終点ハメ, 終点cav, 端末leftRAN)
        
        Dim i As Long
        Dim outdats(1 To 38) As Variant

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=UTF-8" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/" & 構成 & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & 構成 & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>構成:" & 構成 & " " & 色呼 & "</td><td>" & 端末from & " to " & 端末to & "</td><td>SUB:" & サブ & "</td><td>Ver:" & myVer & "</td>" & _
                               "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                               "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
        '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        Dim 端末fromleft As Single, 端末toleft As Single, 端末from1 As String, 端末from2 As String, 端末to1 As String, 端末to2 As String
        端末fromleft = 0: 端末toleft = 0
        For i = LBound(端末leftRAN, 2) + 1 To UBound(端末leftRAN, 2)
            If 端末from = 端末leftRAN(0, i) Then 端末fromleft = 端末leftRAN(1, i)
            If 端末to = 端末leftRAN(0, i) Then 端末toleft = 端末leftRAN(1, i)
        Next i
        '右にある方を右に表示させるbox6と7だと右になる
        If val(端末fromleft) >= val(端末toleft) Then
            端末from1 = "box6"
            端末from2 = "box7"
            端末to1 = "box4"
            端末to2 = "box5"
        Else
            端末from1 = "box4"
            端末from2 = "box5"
            端末to1 = "box6"
            端末to2 = "box7"
        End If
        
        '2.191.01
        If Left(始点ハメ, 1) = "後" Or 先ハメ点滅 = True Then
            outdats(14) = "  <div  id=" & Chr(34) & 端末from1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_1.png" & Chr(34) & " ></div>"
            outdats(15) = "  <div  id=" & Chr(34) & 端末from2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_1_" & 始点cav & ".png" & Chr(34) & " ></div>"
        End If
        If Left(終点ハメ, 1) = "後" Or 先ハメ点滅 = True Then
            outdats(16) = "  <div id=" & Chr(34) & 端末to1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末to & "_1.png" & Chr(34) & " ></div>"
            outdats(17) = "  <div id=" & Chr(34) & 端末to2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末to & "_1_" & 終点cav & ".png" & Chr(34) & " ></div>"
        End If
        outdats(18) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & サブ2 & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"
        outdats(19) = "  <div id=" & Chr(34) & "box2" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "to" & 端末to & "_" & 色呼 & ".png" & Chr(34) & " ></div>"
        outdats(20) = "  <div id=" & Chr(34) & "box3" & Chr(34) & "><img src=" & Chr(34) & "./img/" & サブ & "_foot.png" & Chr(34) & " ></div>"
        outdats(21) = "</body>"
        
        outdats(22) = "<!-- <script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(23) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink2.js" & Chr(34) & "></script> -->"
        outdats(24) = "<script>"
        outdats(25) = "function checkText(){"
        outdats(26) = "  var str1=document.myform.txtb.value;"
        outdats(27) = "  var seihin,kosei;"
        outdats(28) = "  var myLen=str1.length;"
        outdats(29) = "  if (myLen <=10){"
        outdats(30) = "    kosei=str1;"
        outdats(31) = "  }else{"
        outdats(32) = "    seihin=str1.substr(25,10);"
        outdats(33) = "    kosei=str1.substr(11,4);"
        outdats(34) = "  }"
        outdats(35) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(36) = "}"
        outdats(37) = "</script>"
        outdats(38) = "</html>"

        Dim txtFile As String
        txtFile = myDir
        Dim adoSt As ADODB.Stream
        Set adoSt = New ADODB.Stream
        
        Dim strLine As String
        
        With adoSt
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            For i = LBound(outdats) To UBound(outdats)
                strLine = outdats(i)
                .WriteText strLine, adWriteLine
            Next i
            
            'ここからBOM無しにする処理
            .position = 0
            .Type = adTypeBinary
            .position = 3 'BOMデータは3バイト目まで
            Dim byteData() As Byte '一時格納
            byteData = .Read  '一時格納用変数に保存
            .Close 'ストリームを閉じる_リセット
            .Open
            .Write byteData
            .SaveToFile txtFile, adSaveCreateOverWrite
            .Close
        End With
End Function
Public Function TEXT出力_後ハメ点滅html_UTF8(myDir, 端末, sourcePath)
        
        Dim i As Long
        Dim outdats(1 To 38) As Variant

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=UTF-8" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/" & 構成 & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & 構成 & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>構成:" & 構成 & " " & 色呼 & "</td><td>" & 端末from & " to " & 端末to & "</td><td>SUB:" & サブ & "</td><td>Ver:" & myVer & "</td>" & _
                               "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                               "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
        '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        Dim 端末fromleft As Single, 端末toleft As Single, 端末from1 As String, 端末from2 As String, 端末to1 As String, 端末to2 As String
        端末fromleft = 0: 端末toleft = 0
        For i = LBound(端末leftRAN, 2) + 1 To UBound(端末leftRAN, 2)
            If 端末from = 端末leftRAN(0, i) Then 端末fromleft = 端末leftRAN(1, i)
            If 端末to = 端末leftRAN(0, i) Then 端末toleft = 端末leftRAN(1, i)
        Next i
        '右にある方を右に表示させるbox6と7だと右になる
        If val(端末fromleft) >= val(端末toleft) Then
            端末from1 = "box6"
            端末from2 = "box7"
            端末to1 = "box4"
            端末to2 = "box5"
        Else
            端末from1 = "box4"
            端末from2 = "box5"
            端末to1 = "box6"
            端末to2 = "box7"
        End If
        
        '2.191.01
        If Left(始点ハメ, 1) = "後" Or 先ハメ点滅 = True Then
            outdats(14) = "  <div  id=" & Chr(34) & 端末from1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_1.png" & Chr(34) & " ></div>"
            outdats(15) = "  <div  id=" & Chr(34) & 端末from2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_1_" & 始点cav & ".png" & Chr(34) & " ></div>"
        End If
        If Left(終点ハメ, 1) = "後" Or 先ハメ点滅 = True Then
            outdats(16) = "  <div id=" & Chr(34) & 端末to1 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末to & "_1.png" & Chr(34) & " ></div>"
            outdats(17) = "  <div id=" & Chr(34) & 端末to2 & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末to & "_1_" & 終点cav & ".png" & Chr(34) & " ></div>"
        End If
        outdats(18) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & サブ2 & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"
        outdats(19) = "  <div id=" & Chr(34) & "box2" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "to" & 端末to & "_" & 色呼 & ".png" & Chr(34) & " ></div>"
        outdats(20) = "  <div id=" & Chr(34) & "box3" & Chr(34) & "><img src=" & Chr(34) & "./img/" & サブ & "_foot.png" & Chr(34) & " ></div>"
        outdats(21) = "</body>"
        
        outdats(22) = "<!-- <script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(23) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink2.js" & Chr(34) & "></script> -->"
        outdats(24) = "<script>"
        outdats(25) = "function checkText(){"
        outdats(26) = "  var str1=document.myform.txtb.value;"
        outdats(27) = "  var seihin,kosei;"
        outdats(28) = "  var myLen=str1.length;"
        outdats(29) = "  if (myLen <=10){"
        outdats(30) = "    kosei=str1;"
        outdats(31) = "  }else{"
        outdats(32) = "    seihin=str1.substr(25,10);"
        outdats(33) = "    kosei=str1.substr(11,4);"
        outdats(34) = "  }"
        outdats(35) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(36) = "}"
        outdats(37) = "</script>"
        outdats(38) = "</html>"

        Dim txtFile As String
        txtFile = myDir
        Dim adoSt As ADODB.Stream
        Set adoSt = New ADODB.Stream
        
        Dim strLine As String
        
        With adoSt
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            For i = LBound(outdats) To UBound(outdats)
                strLine = outdats(i)
                .WriteText strLine, adWriteLine
            Next i
            
            'ここからBOM無しにする処理
            .position = 0
            .Type = adTypeBinary
            .position = 3 'BOMデータは3バイト目まで
            Dim byteData() As Byte '一時格納
            byteData = .Read  '一時格納用変数に保存
            .Close 'ストリームを閉じる_リセット
            .Open
            .Write byteData
            .SaveToFile txtFile, adSaveCreateOverWrite
            .Close
        End With
End Function




Public Function TEXT出力_配索経路_端末経路html(myDir, 端末from, 端末to, 製品品番str, サブ, 構成, 色呼)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 34) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir
    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift-jis" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/tanmatukeiro.css" & Chr(34) & ">"
        outdats(6) = "<title>" & 端末from & "-</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "myBlink();document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>端末: " & 端末from & "-</td><td>" & 色呼 & "</td><td>" & サブ & "</td><td>Ver:" & myVer & "</td>" & _
                                "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                                "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(14) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & サブ & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"
        outdats(15) = "  <div id=" & Chr(34) & "box4" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_2" & "_foot.png" & Chr(34) & " ></div>"
        outdats(16) = "  <div id=" & Chr(34) & "box2" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_2" & ".png" & Chr(34) & " ></div>"
        outdats(17) = "  <div id=" & Chr(34) & "box3" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_2" & "_tansen.png" & Chr(34) & " ></div>"
        outdats(18) = "</body>"
        
        outdats(19) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(20) = "<script>"
        outdats(21) = "function checkText(){"
        outdats(22) = "  var str1=document.myform.txtb.value;"
        outdats(23) = "  var seihin,kosei;"
        outdats(24) = "  var myLen=str1.length;"
        outdats(25) = "  if (myLen <=10){"
        outdats(26) = "    kosei=str1;"
        outdats(27) = "  }else{"
        outdats(28) = "    seihin=str1.substr(25,10);"
        outdats(29) = "    kosei=str1.substr(11,4);"
        outdats(30) = "  }"
        outdats(31) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(32) = "}"
        outdats(33) = "</script>"
        outdats(34) = "</html>"
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function

Public Function TEXT出力_配索経路_端末経路html_UTF8(myDir, 端末from, 端末to, 製品品番str, サブ, 構成, 色呼)
    
    Dim i As Integer
    Dim outdats(1 To 34) As Variant

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=UTF-8" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/tanmatukeiro.css" & Chr(34) & ">"
        outdats(6) = "<title>" & 端末from & "-</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "myBlink();document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>端末: " & 端末from & "-</td><td>" & 色呼 & "</td><td>" & サブ & "</td><td>Ver:" & myVer & "</td>" & _
                                "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                                "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(14) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & サブ & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"
        outdats(15) = "  <div id=" & Chr(34) & "box4" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_2" & "_foot.png" & Chr(34) & " ></div>"
        outdats(16) = "  <div id=" & Chr(34) & "box2" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_2" & ".png" & Chr(34) & " ></div>"
        outdats(17) = "  <div id=" & Chr(34) & "box3" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末from & "_2" & "_tansen.png" & Chr(34) & " ></div>"
        outdats(18) = "</body>"
        
        outdats(19) = "<script type = " & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "myBlink.js" & Chr(34) & "></script>"
        outdats(20) = "<script>"
        outdats(21) = "function checkText(){"
        outdats(22) = "  var str1=document.myform.txtb.value;"
        outdats(23) = "  var seihin,kosei;"
        outdats(24) = "  var myLen=str1.length;"
        outdats(25) = "  if (myLen <=10){"
        outdats(26) = "    kosei=str1;"
        outdats(27) = "  }else{"
        outdats(28) = "    seihin=str1.substr(25,10);"
        outdats(29) = "    kosei=str1.substr(11,4);"
        outdats(30) = "  }"
        outdats(31) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(32) = "}"
        outdats(33) = "</script>"
        outdats(34) = "</html>"
        
        Dim txtFile As String
        txtFile = myDir
        Dim adoSt As ADODB.Stream
        Set adoSt = New ADODB.Stream
        Dim strLine As String
        
        With adoSt
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            For i = LBound(outdats) To UBound(outdats)
                strLine = outdats(i)
                .WriteText strLine, adWriteLine
            Next i
            'ここからBOM無しにする処理
            .position = 0
            .Type = adTypeBinary
            .position = 3 'BOMデータは3バイト目まで
            Dim byteData() As Byte '一時格納
            byteData = .Read  '一時格納用変数に保存
            .Close 'ストリームを閉じる_リセット
            .Open
            .Write byteData
            .SaveToFile txtFile, adSaveCreateOverWrite
            .Close
        End With

End Function


Public Function TEXT出力_配索経路_端末html(myDir, 端末str, 端末0, 部品品番str)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 31) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    myPath = myDir
    端末0 = "端末:" & 端末0
    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=Shift_JIS" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/" & "atohame" & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & point & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>" & 端末0 & "</td><td>" & 部品品番str & "</td><td>" & "" & "</td>" _
                               & "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                               "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(14) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末str & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"

        outdats(15) = "</body>"
        
        outdats(16) = "<script>"
        outdats(17) = "function checkText(){"
        outdats(18) = "  var str1=document.myform.txtb.value;"
        outdats(19) = "  var seihin,kosei;"
        outdats(20) = "  var myLen=str1.length;"
        outdats(21) = "  if (myLen <=10){"
        outdats(22) = "    kosei=str1;"
        outdats(23) = "  }else{"
        outdats(24) = "    seihin=str1.substr(25,10);"
        outdats(25) = "    kosei=str1.substr(11,4);"
        outdats(26) = "  }"
        outdats(27) = "  "
        outdats(28) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(29) = "}"
        outdats(30) = "</script>"
        outdats(31) = "</html>"
        
        '配列の要素を結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function

Public Function TEXT出力_配索経路_端末html_UTF8(myDir, 端末str, 端末0, 部品品番str)
    
    Dim i As Integer
    Dim outdats(1 To 31) As Variant

    端末0 = "端末:" & 端末0

        outdats(1) = "<html>"
        outdats(2) = "<head>"
        outdats(3) = "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=UTF-8" & Chr(34) & ">"
        outdats(4) = "<meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=11" & Chr(34) & " />"
        outdats(5) = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " media=" & Chr(34) & "all" & Chr(34) & " href=" & Chr(34) & "./css/" & "atohame" & ".css" & Chr(34) & ">"
        outdats(6) = "<title>" & 端末0 & "</title>"
        outdats(7) = "</head>"
        
        outdats(8) = "<body onLoad=" & Chr(34) & "document.myform.txtb.focus();" & Chr(34) & ">"
        
        outdats(9) = "<table>"
        outdats(10) = "<form name=" & Chr(34) & "myform" & Chr(34) & " onsubmit=" & Chr(34) & "return checkText()" & Chr(34) & ">"
        outdats(11) = "<tr class=" & Chr(34) & "top" & Chr(34) & "><td>" & 端末0 & "</td><td>" & 部品品番str & "</td><td>" & "" & "</td><td>" & myVer & "</td>" _
                               & "<td><input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "txtb" & Chr(34) & " onfocus=" & Chr(34) & "this.select()" & Chr(34) & "/>" & _
                               "<input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Enter" & Chr(34) & " class=" & Chr(34) & "myB" & Chr(34) & "></td></tr>"
        outdats(12) = "</from>"
        outdats(13) = "</table>"
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(14) = "  <div class=" & Chr(34) & "box1" & Chr(34) & "><img src=" & Chr(34) & "./img/" & 端末str & ".png" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " /></div>"

        outdats(15) = "</body>"
        
        outdats(16) = "<script>"
        outdats(17) = "function checkText(){"
        outdats(18) = "  var str1=document.myform.txtb.value;"
        outdats(19) = "  var seihin,kosei;"
        outdats(20) = "  var myLen=str1.length;"
        outdats(21) = "  if (myLen <=10){"
        outdats(22) = "    kosei=str1;"
        outdats(23) = "  }else{"
        outdats(24) = "    seihin=str1.substr(25,10);"
        outdats(25) = "    kosei=str1.substr(11,4);"
        outdats(26) = "  }"
        outdats(27) = "  "
        outdats(28) = "  document.myform.action = " & Chr(34) & Chr(34) & "+kosei+" & Chr(34) & ".html" & Chr(34) & ";"
        outdats(29) = "}"
        outdats(30) = "</script>"
        outdats(31) = "</html>"
        
        Dim txtFile As String
        txtFile = myDir
        Dim adoSt As ADODB.Stream
        Set adoSt = New ADODB.Stream
        Dim strLine As String
        
        With adoSt
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            For i = LBound(outdats) To UBound(outdats)
                strLine = outdats(i)
                .WriteText strLine, adWriteLine
            Next i
            'ここからBOM無しにする処理
            .position = 0
            .Type = adTypeBinary
            .position = 3 'BOMデータは3バイト目まで
            Dim byteData() As Byte '一時格納
            byteData = .Read  '一時格納用変数に保存
            .Close 'ストリームを閉じる_リセット
            .Open
            .Write byteData
            .SaveToFile txtFile, adSaveCreateOverWrite
            .Close
        End With

End Function

Public Function TEXT出力_先ハメ誘導_SCC無し_html_UTF8(myDir, 端末1, 端末2, 部品品番1, 部品品番2, subStr, number, kNumbers, pTerm)

    Dim w As String, S As String, html_ary As Variant
    w = Chr(34): S = Chr(39)
    
    Dim num As Variant, Knumber As Variant, i As Long, kNumberText As String, colorCode As Variant, myDiv As String, colorCode1 As String, colorCode2 As String
    num = Split(number, "_")
    kNumberText = ""
    Knumber = Split(kNumbers, ",")
    For i = LBound(Knumber) To UBound(Knumber)
        colorCode = Split(Knumber(i), ";")
        Call exChangeHTMLcolor(colorCode(1), colorCode1, colorCode2, "")
        If InStr(colorCode(1), "/") = 0 Then
            myDiv = "<div " & "class = " & w & "wireColorBack" & w & " style=" & w & "background-color:#" & colorCode1 & ";" & w & "></div>" & vbLf
        Else
            myDiv = "<div class = " & w & "wireColorFront" & w & " style=" & w & "background-color:#" & colorCode2 & ";" & w & "></div>" & vbLf & _
                          "<div " & "class = " & w & "wireColorBack" & w & " style=" & w & "background-color:#" & colorCode1 & ";" & w & "></div>" & vbLf
        End If
        kNumberText = kNumberText & myDiv & vbLf & colorCode(0) & "<br>"
    Next i
    'カウント本数
    kNumberText = kNumberText & UBound(Knumber) + 1 & "<br>"
    
    
    html_ary = readText_UTF8(myAddress(0, 1) & "\41\0000.html")
    
    Dim progressValue As Single
    progressValue = val(num(0)) / val(num(1)) * 100
    
    progressString = "<div id=" & w & "ProgressBar" & w & " style=" & w & "width:" & progressValue & "%;height:5px;background-color:#55c500;" & w & "></div>"
    '親の端末ナンバーと同じならマーク
    Dim pString As String
    If 端末1 = pTerm Then pString = "<span id=" & w & "g" & w & ">★</span> "
    
    Page_ = num(0) & "/" & num(1)
    
    If num(0) <> num(1) Then
        Top_ = "<td onclick=" & w & "pageMove(" & S & "37" & S & ");" & w & ">" & pString & 端末1 & "_" & 部品品番1 & "</td><td id=" & w & "sub" & w & ">SUB:" & subStr & "</td><td onclick=" & w & "pageMove(" & S & "39" & S & ");" & w & ">" & 端末2 & "_" & 部品品番2 & "</td>"""
    Else
        Top_ = "<td onclick=" & w & "pageMove(" & S & "37" & S & ");" & w & ">" & pString & 端末1 & "_" & 部品品番1 & "</td><td id=" & w & "sub" & w & ">SUB:" & subStr & "</td><td>" & 端末2 & "_" & 部品品番2 & "</td>"""
    End If
    If 端末2 <> "" Then
        left_ = "<div id=" & w & "box2" & w & "><img src=" & w & "./img/" & 端末1 & "-" & 端末2 & ".png" & w & "></div><div class=" & w & "box1" & w & "><img src=" & w & "./img/" & 端末1 & ".png" & w & "></div>"
        right_ = "<div id=" & w & "box4" & w & "><img src=" & w & "./img/" & 端末2 & "-" & 端末1 & ".png" & w & "></div><div class=" & w & "box3" & w & "><img src=" & w & "./img/" & 端末2 & ".png" & w & "></div>"
    Else
        left_ = "<div id=" & w & "box2" & w & "><img src=" & w & "./img/" & 端末1 & "-" & "0" & ".png" & w & "></div><div class=" & w & "box1" & w & "><img src=" & w & "./img/" & 端末1 & ".png" & w & "></div>"
        right_ = ""
    End If
    
    '置き換えていく
    html_ary = Replace(html_ary, "__wiretext__", kNumberText)
    html_ary = Replace(html_ary, "__progressbar__", progressString)
    html_ary = Replace(html_ary, "__page__", Page_)
    html_ary = Replace(html_ary, "__top__", Top_)
    html_ary = Replace(html_ary, "__left__", left_)
    html_ary = Replace(html_ary, "__right__", right_)
    
    exportHTML_navi_UTF8 html_ary, myDir
    
End Function


Public Function TEXT出力_先ハメ誘導_SCC無し_single_html_UTF8(myDir, 端末1, 部品品番1, subStr, number, kNumbers, pTerm)
    
    Dim w As String, S As String
    w = Chr(34): S = Chr(39)
        
    Dim num As Variant, Knumber As Variant, i As Long, kNumberText As String, colorCode As Variant, myDiv As String, colorCode1 As String, colorCode2 As String
    num = Split(number, "_")
    Knumber = Split(kNumbers, ",")
    For i = LBound(Knumber) To UBound(Knumber)
        colorCode = Split(Knumber(i), ";")
        Call exChangeHTMLcolor(colorCode(1), colorCode1, colorCode2, "")
        If InStr(colorCode(1), "/") = 0 Then
            myDiv = "<div style=" & w & "background-color:#" & colorCode1 & ";height:1em;width:1em;float:left;position:relative;border:1px solid #111111;top:-1px;left:-1px;" & w & "></div>" & vbLf
        Else
            myDiv = "<div style=" & w & "background-color:#" & colorCode1 & ";height:1em;width:1em;float:left;position:relative;border:1px solid #111111;top:-1px;left:-1px;" & w & "></div>" & vbLf & _
                          "<div style=" & w & "background-color:#" & colorCode2 & ";height:1em;width:0.4em;float:left;position:absolute;left:1.3em;" & w & "></div>" & vbLf
        End If
        kNumberText = kNumberText & myDiv & vbLf & colorCode(0) & "<br>"
    Next i
    'カウント本数
    kNumberText = kNumberText & UBound(Knumber) + 1 & "<br>"
    
    Dim progressValue As Single
    progressValue = val(num(0)) / val(num(1)) * 100
    
    '親の端末ナンバーと同じならマーク
    Dim pString As String
    If 端末1 = pTerm Then
        pString = "<span id=" & w & "g" & w & ">☆</span> "
    End If
    Dim outdats(1 To 23) As Variant

        outdats(1) = "<html>"
        outdats(2) = "    <head>"
        outdats(3) = "        <meta http-equiv=" & w & "content-type" & w & " content=" & w & "text/html; charset=UTF-8" & w & ">"
        outdats(4) = "        <meta http-equiv=" & w & "X-UA-Compatible" & w & " content=" & w & "IE=edge" & w & " />"
        outdats(5) = "        <link rel=" & w & "stylesheet" & w & " type=" & w & "text/css" & w & " media=" & w & "all" & w & " href=" & w & "./css/" & "single.css" & w & ">"
        outdats(6) = "        <title>" & num(0) & "/" & num(1) & "</title>"
        outdats(7) = "    </head>"
    
        outdats(8) = "    <body onLoad=" & w & "myBlink();keyEvents();" & w & ">"
        outdats(9) = "        <table>"
        outdats(10) = "            <form name=" & w & "myform" & w & ">"
        outdats(11) = "               <tr class=" & w & "top" & w & ">" & _
                                               "<td onclick=" & w & "pageMove(" & S & "37" & S & ");" & w & ">" & pString & 端末1 & "_" & 部品品番1 & "</td>" & _
                                               "<td id=" & w & "sub" & w & ">SUB:" & subStr & "</td>" & _
                                               "<td onclick=" & w & "pageMove(" & S & "39" & S & ");" & w & ">" & "" & "" & "" & "</td>"
        outdats(12) = "            </from>"
        outdats(13) = "        </table>"
        outdats(14) = "        <div id=" & w & "ProgressBar" & w & " style=" & w & "width:" & progressValue & "%;" & w & "></div>"
        
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(15) = "        <div id =" & w & "right" & w & " onclick=" & w & "pageMove(" & S & "39" & S & ");" & w & ">"
        outdats(16) = "            <div class=" & w & "box1" & w & "><img src=" & w & "./img/" & 端末1 & ".png" & w & "></div>"
        outdats(17) = "            <div id=" & w & "box2" & w & "><img src=" & w & "./img/" & 端末1 & "-0" & ".png" & w & "></div>"
        outdats(18) = "        </div>"
        outdats(19) = "        <div id=" & w & "wireText" & w & ">" & kNumberText & "</div>"
        outdats(20) = "    </body>"
        
        outdats(21) = "    <script type =" & w & "text/javascript" & w & " src=" & w & "myBlink_single.js" & w & "></script>"
        outdats(22) = "    <script type =" & w & "text/javascript" & w & " src=" & w & "keyEvents.js" & w & "></script>"
        
        outdats(23) = "</html>"
        
        Dim txtFile As String
        txtFile = myDir
        Dim adoSt As ADODB.Stream
        Set adoSt = New ADODB.Stream
        Dim strLine As String
        
        With adoSt
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            For i = LBound(outdats) To UBound(outdats)
                strLine = outdats(i)
                .WriteText strLine, adWriteLine
            Next i
            'ここからBOM無しにする処理
            .position = 0
            .Type = adTypeBinary
            .position = 3 'BOMデータは3バイト目まで
            Dim byteData() As Byte '一時格納
            byteData = .Read  '一時格納用変数に保存
            .Close 'ストリームを閉じる_リセット
            .Open
            .Write byteData
            .SaveToFile txtFile, adSaveCreateOverWrite
            .Close
        End With

End Function

Public Function TEXT出力_先ハメ誘導_SCC無し_double_html_UTF8(myDir, 端末1, 端末2, 部品品番1, 部品品番2, subStr, number, kNumbers, pTerm)

    Dim w As String, S As String
    w = Chr(34): S = Chr(39)
        
    Dim num As Variant, Knumber As Variant, i As Long, kNumberText As String, colorCode As Variant, myDiv As String, colorCode1 As String, colorCode2 As String
    num = Split(number, "_")
    Knumber = Split(kNumbers, ",")
    For i = LBound(Knumber) To UBound(Knumber)
        colorCode = Split(Knumber(i), ";")
        Call exChangeHTMLcolor(colorCode(1), colorCode1, colorCode2, "")
        If InStr(colorCode(1), "/") = 0 Then
            myDiv = "<div style=" & w & "background-color:#" & colorCode1 & ";height:1em;width:1em;float:left;position:relative;border:1px solid #111111;top:-1px;left:-1px;" & w & "></div>" & vbLf
        Else
            myDiv = "<div style=" & w & "background-color:#" & colorCode1 & ";height:1em;width:1em;float:left;position:relative;border:1px solid #111111;top:-1px;left:-1px;" & w & "></div>" & vbLf & _
                          "<div style=" & w & "background-color:#" & colorCode2 & ";height:1em;width:0.4em;float:left;position:absolute;left:1.3em;" & w & "></div>" & vbLf
        End If
        kNumberText = kNumberText & myDiv & vbLf & colorCode(0) & "<br>"
    Next i
    'カウント本数
    kNumberText = kNumberText & UBound(Knumber) + 1 & "<br>"
    
    Dim progressValue As Single
    progressValue = val(num(0)) / val(num(1)) * 100
    '親の端末ナンバーと同じならマーク
    Dim pString As String
    If 端末1 = pTerm Then pString = "<span id=" & w & "g" & w & ">☆</span> "
    
    Dim outdats(1 To 27) As Variant
        
        outdats(1) = "<html>"
        outdats(2) = "    <head>"
        outdats(3) = "        <meta http-equiv=" & w & "content-type" & w & " content=" & w & "text/html; charset=UTF-8" & w & ">"
        outdats(4) = "        <meta http-equiv=" & w & "X-UA-Compatible" & w & " content=" & w & "IE=edge" & w & " />"
        outdats(5) = "        <link rel=" & w & "stylesheet" & w & " type=" & w & "text/css" & w & " media=" & w & "all" & w & " href=" & w & "./css/" & "double.css" & w & ">"
        outdats(6) = "        <title>" & num(0) & "/" & num(1) & "</title>"
        outdats(7) = "    </head>"
        
        outdats(8) = "    <body onLoad=" & w & "myBlink();keyEvents();" & w & ">"
        outdats(9) = "        <table>"
        outdats(10) = "            <form name=" & w & "myform" & w & ">"
        outdats(11) = "                <tr class=" & w & "top" & w & ">" & _
                                                 "<td onclick=" & w & "pageMove(" & S & "37" & S & ");" & w & ">" & pString & 端末1 & "_" & 部品品番1 & "</td>" & _
                                                 "<td id=" & w & "sub" & w & ">SUB:" & subStr & "</td>" & _
                                                 "<td onclick=" & w & "pageMove(" & S & "39" & S & ");" & w & ">" & 端末2 & "_" & 部品品番2 & "</td>"
        outdats(12) = "            </from>"
        outdats(13) = "        </table>"
        outdats(14) = "        <div id=" & w & "ProgressBar" & w & " style=" & w & "width:" & progressValue & "%;" & w & "></div>"
                '<div style="position:absolute; top:0px; left:0px;"><img src="Base.png" width="1220" height="480" alt="" border="0"></div>
        outdats(15) = "        <div id =" & w & "left" & w & " onclick=" & w & "pageMove(" & S & "37" & S & ");" & w & ">"
        outdats(16) = "            <div class=" & w & "box1" & w & "><img src=" & w & "./img/" & 端末1 & ".png" & w & "></div>"
        outdats(17) = "            <div id=" & w & "box2" & w & "><img src=" & w & "./img/" & 端末1 & "-" & 端末2 & ".png" & w & "></div>"
        outdats(18) = "        </div>"
        outdats(19) = "        <div id =" & w & "right" & w & " onclick=" & w & "pageMove(" & S & "39" & S & ");" & w & ">"
        outdats(20) = "            <div class=" & w & "box3" & w & "><img src=" & w & "./img/" & 端末2 & ".png" & w & "></div>"
        outdats(21) = "            <div id=" & w & "box4" & w & "><img src=" & w & "./img/" & 端末2 & "-" & 端末1 & ".png" & w & "></div>"
        outdats(22) = "        </div>"
        outdats(23) = "        <div id=" & w & "wireText" & w & ">" & kNumberText & "</div>"
        outdats(24) = "    </body>"
        
        outdats(25) = "    <script type =" & w & "text/javascript" & w & " src=" & w & "myBlink_double.js" & w & "></script>"
        outdats(26) = "    <script type =" & w & "text/javascript" & w & " src=" & w & "keyEvents.js" & w & "></script>"
        outdats(27) = "</html>"
        
        Dim txtFile As String
        txtFile = myDir
        Dim adoSt As ADODB.Stream
        Set adoSt = New ADODB.Stream
        Dim strLine As String
        
        With adoSt
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            For i = LBound(outdats) To UBound(outdats)
                strLine = outdats(i)
                .WriteText strLine, adWriteLine
            Next i
            'ここからBOM無しにする処理
            .position = 0
            .Type = adTypeBinary
            .position = 3 'BOMデータは3バイト目まで
            Dim byteData() As Byte '一時格納
            byteData = .Read  '一時格納用変数に保存
            .Close 'ストリームを閉じる_リセット
            .Open
            .Write byteData
            .SaveToFile txtFile, adSaveCreateOverWrite
            .Close
        End With

End Function


Public Function TEXT出力_設定_竿レイアウト図(myDir)
    
    Dim myPath          As String
    Dim FileNumber      As Integer
    Dim outdats(1 To 4) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    mess0 = "2行目にCav穴番号の変換に使用するファイル名、3行目に部材一覧のディレクトリを入力してください。"
    mess1 = Left(myDir, InStr(myDir, "生産準備+") + 4) & "\010_手入力情報\Exchange_CavToHole.xlsx"
    mess2 = myAddress(1, 1)
    mess3 = Left(myDir, InStr(myDir, "生産準備+") + 4) & "\010_手入力情報\自動機設定.xlsx"
    
    myPath = myDir
    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber
        
        outdats(1) = mess0
        outdats(2) = mess1
        outdats(3) = mess2
        outdats(4) = mess3
        
        '配列の要素を結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function


Public Function TEXT出力_汎用検査履歴システムjs(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 7) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber
    
        outdats(1) = " mSec = 300; //  点滅速度 (1sec=1000)"
        outdats(2) = " function myBlink(){"
        outdats(3) = "     flag = document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility;"
        outdats(4) = "     if (flag == " & Chr(34) & "visible" & Chr(34) & ") document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";"
        outdats(5) = "     else document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";"
        outdats(6) = "     setTimeout(" & Chr(34) & "myBlink()" & Chr(34) & ",mSec);"
        outdats(7) = " }"
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function

Public Function TEXT出力_配索経路_端末js(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 20) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber
    
        outdats(1) = " //mSec = 300; //  点滅速度 (1sec=1000)"
        outdats(2) = " function myBlink(){"
        outdats(3) = "     flag = document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility;"
        outdats(4) = "     if (flag == " & Chr(34) & "visible" & Chr(34) & "){"
        outdats(5) = "         document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";"
        outdats(6) = "         mSec = 600;"
        outdats(7) = "     }else {"
        outdats(8) = "         document.getElementById(" & Chr(34) & "box2" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";"
        outdats(9) = "         mSec = 300;"
        outdats(10) = "     }"
        outdats(11) = "     flag = document.getElementById(" & Chr(34) & "box3" & Chr(34) & ").style.visibility;"
        outdats(12) = "     if (flag == " & Chr(34) & "hidden" & Chr(34) & "){"
        outdats(13) = "         document.getElementById(" & Chr(34) & "box3" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";"
        outdats(14) = "         mSec = 600;"
        outdats(15) = "     }else {"
        outdats(16) = "         document.getElementById(" & Chr(34) & "box3" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";"
        outdats(17) = "         mSec = 300;"
        outdats(18) = "     }"
        outdats(19) = "     setTimeout(" & Chr(34) & "myBlink()" & Chr(34) & ",mSec);"
        outdats(20) = " }"
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function
Public Function TEXT出力_配索経路_端末js2(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 17) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber
    
        outdats(1) = " mSec = 300; //  点滅速度 (1sec=1000)"
        outdats(2) = " function myBlink2(){"
        outdats(3) = " mSec = 175;"
        outdats(4) = "     try{flag = document.getElementById(" & Chr(34) & "box5" & Chr(34) & ").style.visibility;} catch(e){}"
        outdats(5) = "     if (flag == " & Chr(34) & "visible" & Chr(34) & "){"
        outdats(6) = "         try{document.getElementById(" & Chr(34) & "box5" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";} catch(e){}"
        outdats(7) = "     }else {"
        outdats(8) = "         try{document.getElementById(" & Chr(34) & "box5" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";} catch(e){}"
        outdats(9) = "     }"
        outdats(10) = "     try{flag = document.getElementById(" & Chr(34) & "box7" & Chr(34) & ").style.visibility;} catch(e){}"
        outdats(11) = "     if (flag == " & Chr(34) & "visible" & Chr(34) & "){"
        outdats(12) = "         try{document.getElementById(" & Chr(34) & "box7" & Chr(34) & ").style.visibility = " & Chr(34) & "hidden" & Chr(34) & ";} catch(e){}"
        outdats(13) = "     }else {"
        outdats(14) = "         try{document.getElementById(" & Chr(34) & "box7" & Chr(34) & ").style.visibility = " & Chr(34) & "visible" & Chr(34) & ";} catch(e){}"
        outdats(15) = "     }"
        outdats(16) = "     setTimeout(" & Chr(34) & "myBlink2()" & Chr(34) & ",mSec);"
        outdats(17) = " }"
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function

Public Function TEXT出力_配索経路_ver(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 4) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber
    
    outdats(1) = "日付:" & Date
    outdats(2) = "ver:" & myVer
    outdats(3) = "後ハメのみ:" & 配索図作成temp
    outdats(4) = "noPictureCount:" & noPictureCount
    
    Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function

Public Function TEXT出力_配索経路css(myPath, box2l, box2t, box2w, box2h, clocode1, clofont)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 66) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber
    
        outdats(1) = "table {"
        outdats(2) = "    table-layout: fixed;"
        outdats(3) = "    width:100%;"
        outdats(4) = "    background-color:#" & clocode1 & ";"
        outdats(5) = "    text-align:center;"
        outdats(6) = "    color: #" & clofont & ";"
        outdats(7) = "    font-size:14pt;"
        outdats(8) = "    font-weight: bold;"
        outdats(9) = "    border-collapse: collapse;"
        outdats(10) = "    font-family: Verdana, Arial, Helvetica, sans-serif;"
        outdats(11) = "    border-bottom:0px solid #000000;"
        outdats(12) = "}"
        outdats(13) = "table td {"
        outdats(14) = "    border: 1px solid  #" & clofont & "; /* 表内側の線：1px,実線,緑色 */"
        outdats(15) = "    border-left:2px solid #" & clofont & ";"
        outdats(16) = "    border-right:2px solid  #" & clofont & ";"
        outdats(17) = "    padding: 1px;            /* セル内側の余白：3ピクセル */"
        outdats(18) = "}"
        outdats(19) = ".box1 img{"
        outdats(21) = "    position:absolute;"
        outdats(22) = "    width:99%;"
        outdats(23) = "    height:auto;"
        outdats(24) = "    max-width:99%;"
        outdats(25) = "    max-height:95%;"
        outdats(26) = "}"
        outdats(27) = ".box1 {"
        outdats(28) = "}"
        outdats(29) = "#box2 img{"
        outdats(30) = "    filter:alpha(opacity=70); /* IE 6,7*/"
        outdats(31) = "    position: absolute;"
        outdats(32) = "    width:99%;"
        outdats(33) = "    opacity:0.7;"
        outdats(34) = "    zoom:1;"
        outdats(35) = "    display:inline-block;"
        outdats(36) = "}"
        outdats(37) = "#box3 img{"
        outdats(38) = "    position:absolute;"
        outdats(39) = "    width:99%;"
        outdats(40) = "}"
        outdats(41) = "#box4 img{"
        outdats(42) = "    position:absolute;"
        outdats(43) = "    bottom:0%;"
        outdats(44) = "    height:30%;"
        outdats(45) = "}"
        outdats(46) = "#box5 img{"
        outdats(47) = "    position:absolute;"
        outdats(48) = "    bottom:0%;"
        outdats(49) = "    height:30%;"
        outdats(50) = "    filter:alpha(opacity=70);"
        outdats(51) = "    opacity:0.7;"
        outdats(52) = "}"
        outdats(53) = "#box6 img{"
        outdats(54) = "    position:absolute;right:0%;"
        outdats(55) = "    bottom:0%;"
        outdats(56) = "    height:30%;"
        outdats(57) = "}"
        outdats(58) = "#box7 img{"
        outdats(59) = "    position:absolute;right:0%;"
        outdats(60) = "    bottom:0%;"
        outdats(61) = "    height:30%;"
        outdats(62) = "    filter:alpha(opacity=70);"
        outdats(63) = "    opacity:0.7;"
        outdats(64) = "}"
        outdats(65) = "body{background-color:#111111;}"
        outdats(66) = ".myB{color:#" & clofont & ";background-color:#" & clocode1 & ";}"
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function

Public Function TEXT出力_配索経路css_UTF8(myPath, box2l, box2t, box2w, box2h, clocode1, clofont)
    
    Dim outdats(1 To 67) As Variant
    Dim i               As Integer

    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)
    
    outdats(1) = "table {"
    outdats(2) = "    table-layout: fixed;"
    outdats(3) = "    width:100%;"
    outdats(4) = "    background-color:#" & clocode1 & ";"
    outdats(5) = "    text-align:center;"
    outdats(6) = "    color: #" & clofont & ";"
    outdats(7) = "    font-size:14pt;"
    outdats(8) = "    font-weight: bold;"
    outdats(9) = "    border-collapse: collapse;"
    outdats(10) = "    font-family: Verdana, Arial, Helvetica, sans-serif;"
    outdats(11) = "    border-bottom:0px solid #000000;"
    outdats(12) = "}"
    outdats(13) = "table td {"
    outdats(14) = "    border: 1px solid  #" & clofont & "; /* 表内側の線：1px,実線,緑色 */"
    outdats(15) = "    border-left:2px solid #" & clofont & ";"
    outdats(16) = "    border-right:2px solid  #" & clofont & ";"
    outdats(17) = "    padding: 1px;            /* セル内側の余白：3ピクセル */"
    outdats(18) = "}"
    outdats(19) = ".box1 img{"
    outdats(21) = "    position:absolute;"
    outdats(22) = "    width:99%;"
    outdats(23) = "    height:auto;"
    outdats(24) = "    max-width:99%;"
    outdats(25) = "    max-height:95%;"
    outdats(26) = "}"
    outdats(27) = ".box1 {"
    outdats(28) = "}"
    outdats(29) = "#box2 img{"
    outdats(30) = "    filter:alpha(opacity=70); /* IE 6,7*/"
    outdats(31) = "    position: absolute;"
    outdats(32) = "    width:99%;"
    outdats(33) = "    opacity:0.7;"
    outdats(34) = "    zoom:1;"
    outdats(35) = "    display:inline-block;-webkit-animation:blink 0.15s ease-in-out infinite alternate;-moz-animation:blink 0.15s ease-in-out infinite alternate;animation:blink 0.15s ease-in-out infinite alternate;"
    outdats(36) = "}"
    outdats(37) = "#box3 img{"
    outdats(38) = "    position:absolute;"
    outdats(39) = "    width:99%;animation: blink_visible 1s steps(5, start) infinite;-webkit-animation: blink_visible 1s steps(5, start) infinite;animation-delay: 1s;-webkit-animation-delay: 1s;"
    outdats(40) = "}"
    outdats(41) = "#box4 img{"
    outdats(42) = "    position:absolute;"
    outdats(43) = "    bottom:0%;"
    outdats(44) = "    max-width:49%;max-height:30%;height:auto;"
    outdats(45) = "}"
    outdats(46) = "#box5 img{"
    outdats(47) = "    position:absolute;"
    outdats(48) = "    bottom:0%;"
    outdats(49) = "    max-width:49%;max-height:30%;height:auto;"
    outdats(50) = "    filter:alpha(opacity=70);"
    outdats(51) = "    opacity:0.7;-webkit-animation:blink 0.15s ease-in-out infinite alternate;-moz-animation:blink 0.15s ease-in-out infinite alternate;animation:blink 0.15s ease-in-out infinite alternate;"
    outdats(52) = "}"
    outdats(53) = "#box6 img{"
    outdats(54) = "    position:absolute;right:0%;"
    outdats(55) = "    bottom:0%;"
    outdats(56) = "    max-width:49%;max-height:30%;height:auto;"
    outdats(57) = "}"
    outdats(58) = "#box7 img{"
    outdats(59) = "    position:absolute;right:0%;"
    outdats(60) = "    bottom:0%;"
    outdats(61) = "    max-width:49%;max-height:30%;height:auto;"
    outdats(62) = "    filter:alpha(opacity=70);"
    outdats(63) = "    opacity:0.7;-webkit-animation:blink 0.15s ease-in-out infinite alternate;-moz-animation:blink 0.15s ease-in-out infinite alternate;animation:blink 0.15s ease-in-out infinite alternate;"
    outdats(64) = "}"
    outdats(65) = "body{background-color:#111111;}"
    outdats(66) = ".myB{color:#" & clofont & ";background-color:#" & clocode1 & ";}"
    outdats(67) = ".blinking{-webkit-animation:blink 1.5s ease-in-out infinite alternate;-moz-animation:blink 1.5s ease-in-out infinite alternate;animation:blink 1.5s ease-in-out infinite alternate;}@-webkit-keyframes blink{0%{opacity:0;}100% {opacity:1;}}@-moz-keyframes blink{0%{opacity:0;}100%{opacity:1;}}@keyframes blink{0%{opacity:0;}100%{opacity:1;}}@keyframes blink_visible{to {visibility: hidden;}}"
    Dim txtFile As String
    txtFile = myPath
    Dim adoSt As ADODB.Stream
    Set adoSt = New ADODB.Stream
    
    Dim strLine As String
    
    With adoSt
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        For i = LBound(outdats) To UBound(outdats)
            strLine = outdats(i)
            .WriteText strLine, adWriteLine
        Next i
        
        'ここからBOM無しにする処理
        .position = 0
        .Type = adTypeBinary
        .position = 3 'BOMデータは3バイト目まで
        Dim byteData() As Byte '一時格納
        byteData = .Read  '一時格納用変数に保存
        .Close 'ストリームを閉じる_リセット
        .Open
        .Write byteData
        .SaveToFile txtFile, adSaveCreateOverWrite
        .Close
    End With

End Function


Public Function TEXT出力_配索経路_端末css(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 47) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    clocode1 = "FFFFFF"
    clofont = "000000"
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber
    
        outdats(1) = "table {"
        outdats(2) = "    table-layout: fixed;"
        outdats(3) = "    width:100%;"
        outdats(4) = "    background-color:#" & clocode1 & ";"
        outdats(5) = "    text-align:center;"
        outdats(6) = "    color: #" & clofont & ";"
        outdats(7) = "    font-size:14pt;"
        outdats(8) = "    font-weight: bold;"
        outdats(9) = "    border-collapse: collapse;"
        outdats(10) = "    font-family: Verdana, Arial, Helvetica, sans-serif;"
        outdats(11) = "    border-bottom:0px solid #000000;"
        outdats(12) = "}"
        outdats(13) = "table td {"
        outdats(14) = "    border: 1px solid  #" & clofont & "; /* 表内側の線：1px,実線,緑色 */"
        outdats(15) = "    border-left:2px solid #" & clofont & ";"
        outdats(16) = "    border-right:2px solid  #" & clofont & ";"
        outdats(17) = "    padding: 1px;            /* セル内側の余白：3ピクセル */"
        outdats(18) = "}"
        outdats(19) = ".box1 img{"
        outdats(21) = "    position:absolute;"
        outdats(22) = "    width:auto;"
        outdats(23) = "    height:auto;"
        outdats(24) = "    max-width:100%;"
        outdats(25) = "    max-height:95%;"
        outdats(26) = "}"
        outdats(27) = ".box1 {"
        outdats(28) = "}"
        outdats(29) = "#box2 img{"
        outdats(30) = "    filter:alpha(opacity=60); /* IE 6,7*/"
        outdats(31) = "    position: absolute;"
        outdats(32) = "    width:100%;"
        outdats(33) = "    opacity:0.8;"
        outdats(34) = "    zoom:1;"
        outdats(35) = "    display:inline-block;"
        outdats(36) = "}"
        outdats(37) = "#box3 img{"
        outdats(38) = "    position:absolute;"
        outdats(39) = "    width:100%;"
        outdats(40) = "}"
        outdats(41) = "#box4 img{"
        outdats(42) = "    position:absolute;"
        outdats(43) = "    bottom:0%;"
        outdats(44) = "    width:100%;"
        outdats(45) = "}"
        outdats(46) = "body{background-color:#111111;}"
        outdats(47) = ".myB{color:#000000;background-color:#FFFFFF;}"
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function


Public Function TEXT出力_汎用検査履歴システムcss(myPath, clocode1, clofont)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 47) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean
    
    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber
    
        outdats(1) = "table {"
        outdats(2) = "    table-layout: fixed;"
        outdats(3) = "    width:100%;"
        outdats(4) = "    background-color:#" & clocode1 & ";"
        outdats(5) = "    text-align:center;"
        outdats(6) = "    color: #" & clofont & ";"
        outdats(7) = "    font-size:14pt;"
        outdats(8) = "    font-weight: bold;"
        outdats(9) = "    border-collapse: collapse;"
        outdats(10) = "    font-family: Verdana, Arial, Helvetica, sans-serif;"
        outdats(11) = "    border-bottom:0px solid #" & clofont & ";"
        outdats(12) = "}"
        outdats(13) = "table td {"
        outdats(14) = "    border: 1px solid  #" & clofont & "; /* 表内側の線：1px,実線,緑色 */"
        outdats(15) = "    border-left:2px solid #" & clofont & ";"
        outdats(16) = "    border-right:2px solid  #" & clofont & ";"
        outdats(17) = "    padding: 1px;            /* セル内側の余白：3ピクセル */"
        outdats(18) = "}"
        outdats(19) = ".box1 img{"
        outdats(21) = "    position:absolute;"
        outdats(22) = "    width:auto;"
        outdats(23) = "    height:auto;"
        outdats(24) = "    max-width:98%;"
        outdats(25) = "    max-height:95%;"
        outdats(26) = "}"
        outdats(27) = ".box1 {"
        outdats(28) = "}"
        outdats(29) = "#box2 img{"
        outdats(30) = "    filter:alpha(opacity=60); /* IE 6,7*/"
        outdats(31) = "    position: absolute;"
        outdats(32) = "    width:auto;"
        outdats(33) = "    height:auto;"
        outdats(34) = "    max-width:98%;"
        outdats(35) = "    max-height:95%;"
        outdats(36) = "    opacity:0.6;"
        outdats(37) = "    display:inline-block;"
        outdats(38) = "}"
        outdats(39) = "#box3 img{"
        outdats(40) = "    position:absolute;"
        outdats(41) = "    width:100%;"
        outdats(42) = "}"
        outdats(43) = "#box4 img{"
        outdats(44) = "    position:absolute;"
        outdats(45) = "    bottom:0%;"
        outdats(46) = "    width:100%;"
        outdats(47) = "}"
        
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function



Public Function 始点終点入れ替え()
    'フィールド名.rowを超えてなければ処理を出る
    Dim keyRow As Long: keyRow = Cells.Find("電線識別名", , , 1).Row
    If Selection.Row < keyRow Then Exit Function
    
    Call 最適化
    '始点側を含むフィールド名を取得
    Dim changeTitle As String: Dim lastCol As Long
    lastCol = Cells(keyRow, Columns.count).End(xlToLeft).Column
    For x = 1 To lastCol
        If Left(Cells(keyRow, x), 3) = "始点側" Then
            changeTitle = changeTitle & "," & Mid(Cells(keyRow, x), 4)
        End If
    Next x
    '始点側/終点側の列を取得
    Dim gawa(1) As String: gawa(0) = "始点側": gawa(1) = "終点側"
    Dim retsu() As Long
    Dim changeTitleSP As Variant: changeTitleSP = Split(changeTitle, ",")
    ReDim retsu(1, UBound(changeTitleSP))
    For g = 0 To 1
        For u = 1 To UBound(changeTitleSP) '0番目が余分なデータ
            retsu(g, u) = Rows(keyRow).Find(gawa(g) & changeTitleSP(u), , , 1).Column
        Next u
    Next g
    '始点側/終点側を入れ替える
    Dim tempKey As Variant, tempCol As Long
    Set tempKey = Rows(keyRow).Find("終点始点入れ替えtemp", , , 1)
    If tempKey Is Nothing Then
        tempCol = lastCol + 1
        Cells(keyRow, tempCol) = "始点終点入れ替えtemp"
    Else
        tempCol = tempKey.Column
    End If
    'tempの列にコピーして各列毎に始点/終点の入れ替え
    Dim startRow As Long: startRow = Selection.Row
    Dim endRow As Long: endRow = Selection.Row + Selection.Rows.count - 1
    For u = 1 To UBound(changeTitleSP)
        Range(Cells(startRow, retsu(0, u)), Cells(endRow, retsu(0, u))).Copy Destination:=Range(Cells(startRow, tempCol), Cells(endRow, tempCol))
        Range(Cells(startRow, retsu(1, u)), Cells(endRow, retsu(1, u))).Copy Destination:=Range(Cells(startRow, retsu(0, u)), Cells(endRow, retsu(0, u)))
        Range(Cells(startRow, tempCol), Cells(endRow, tempCol)).Copy Destination:=Range(Cells(startRow, retsu(1, u)), Cells(endRow, retsu(1, u)))
    Next u
    '入れ替えた行は列:始終替を1にする
    Dim changeFlgCol As Long: changeFlgCol = Rows(keyRow).Find("始終替", , , 1).Column
    For y = startRow To endRow
        If Cells(y, changeFlgCol) = "1" Then
            Cells(y, changeFlgCol) = Empty
        Else
            Cells(y, changeFlgCol) = "1"
        End If
    Next y
    '解放
    Columns(tempCol).Delete
    Set tempKey = Nothing
    
    Call 最適化もどす
    
    Call PlaySound("けってい")
    
End Function

Public Function 作業色に着色(myNum)
    addressSet ThisWorkbook
    'フィールド名.rowを超えてなければ終了
    Dim keyRow As Long: keyRow = Cells.Find("電線識別名", , , 1).Row
    If Selection.Row < keyRow Then Exit Function

    '[設定]から色を取得
    With Sheets("設定")
        Dim myKey As Range, myRange As Range, myNumF As Range
        Dim myFontColor As Variant, myInteriorColor As Long, myBold As Boolean
        Set myKey = .Cells.Find("ハメ色_", , , 1).Offset(0, 1)
        Set myRange = .Range(myKey, myKey.End(xlDown))
        If myNum = "-" Then
            myFontColor = 0
            myInteriorColor = 16777215
            myBold = False
        Else
            Set myNumF = myRange.Find(myNum, , , 1)
            If myNumF Is Nothing Then Exit Function '呼び出されたmyNumが無ければ終了
            myFontColor = myNumF.Font.color
            myInteriorColor = myNumF.Interior.color
            myBold = True
        End If
    End With

    '始点側または終点側を選択していなければ終了
    Dim selectGawa As String, retsu(2) As Long, 対象名 As String
    selectGawa = Left(Cells(keyRow, Selection.Column), 3)
    If Not selectGawa = "始点側" And Not selectGawa = "終点側" Then Exit Function
    
    '着色する列を取得
    Dim myTitle As String: myTitle = "回路符号,端末識別子,キャビティ"
    Dim myTitleSP As Variant: myTitleSP = Split(myTitle, ",")
    For x = LBound(myTitleSP) To UBound(myTitleSP)
        retsu(x) = Rows(keyRow).Find(selectGawa & myTitleSP(x), , , 1).Column
    Next x

    'tempの列にコピーして各列毎に始点/終点の入れ替え
    Dim startRow As Long: startRow = Selection.Row
    Dim endRow As Long: endRow = Selection.Row + Selection.Rows.count - 1
    For u = LBound(myTitleSP) To UBound(myTitleSP)
        Range(Cells(startRow, retsu(u)), Cells(endRow, retsu(u))).Font.color = myFontColor 'フォント色作業色に着色
        Range(Cells(startRow, retsu(u)), Cells(endRow, retsu(u))).Font.Bold = myBold 'フォントを太字にする

        If myInteriorColor <> 16777215 Then
            Range(Cells(startRow, retsu(u)), Cells(endRow, retsu(u))).Interior.color = myInteriorColor
        Else
            Range(Cells(startRow, retsu(u)), Cells(endRow, retsu(u))).Interior.ColorIndex = xlNone '背景が塗りつぶし無しの時
        End If
    Next u

    '解放
    Set myKey = Nothing
    Set myRange = Nothing
    Set myNumF = Nothing
    
    Call 最適化もどす
    
    Call PlaySound("けってい")
    
End Function

Function 着色_通知書(myNum)
    'フィールド名.rowを超えてなければ終了
    Dim key As Variant
    Set key = Cells.Find("日付_", , , 1)
    If Selection.Row <= key.Row Then Exit Function
    
    Dim myKey As Range, myRange As Range, myNumF As Range
    Dim myFontColor As Variant, myInteriorColor As Long, myBold As Boolean, myString As String
    '[設定]から色を取得
    With Sheets("設定")
        Select Case myNum
            Case "1"
                myFontColor = 0
                myInteriorColor = 5263615
                myString = Date
            Case "2"
                myFontColor = 16777215
                myInteriorColor = 16711680
                myString = Date
            Case "0"
                myFontColor = 16777215
                myInteriorColor = 8421504
                myString = Date
            Case "-"
                myFontColor = 0
                myInteriorColor = -1
                myString = ""
        End Select
    End With
    
    '着色する列を取得
    Dim myTitle As String: myTitle = "PVSW,RLTF,部品リスト,社内図"
    Dim myTitleSP As Variant: myTitleSP = Split(myTitle, ",")
    Dim retsu() As Variant, foundFlag As Boolean
    ReDim retsu(UBound(myTitleSP))
    foundFlag = False
    For x = LBound(myTitleSP) To UBound(myTitleSP)
        retsu(x) = Rows(key.Row).Find(selectGawa & myTitleSP(x), , , 1).Column
        If ActiveCell.Column = retsu(x) Then foundFlag = True
    Next x

    If foundFlag = False Then Exit Function
    
    For Each cell In Selection.Cells
        For x = LBound(retsu) To UBound(retsu)
            If cell.Column = retsu(x) Then
                cell.NumberFormat = "mm/dd"
                cell.Value = myString
                cell.Font.color = myFontColor
                cell.Font.Bold = myBold
                If myInteriorColor = -1 Then
                    cell.Interior.Pattern = xlNone
                Else
                    cell.Interior.color = myInteriorColor
                End If
                Exit For
            End If
        Next x
    Next cell
    
    Dim startCol As Long: startCol = Selection.Column
    Dim endCol As Long: endCol = Selection.Column + Selection.Columns.count - 1
    If startCol < retsu(LBound(myTitleSP)) Then startCol = retsu(LBound(myTitleSP))
    If endCol > retsu(UBound(myTitleSP)) Then endCol = retsu(UBound(myTitleSP))
    Range(Cells(ActiveCell.Row, startCol), Cells(ActiveCell.Row, endCol)).Font.color = myFontColor 'フォント色作業色に着色
    Range(Cells(ActiveCell.Row, startCol), Cells(ActiveCell.Row, endCol)).Font.Bold = myBold 'フォントを太字にする
    If myInteriorColor = -1 Then
        Range(Cells(ActiveCell.Row, startCol), Cells(ActiveCell.Row, endCol)).Interior.ColorIndex = xlNone '背景が塗りつぶし無しの時
    Else
        Range(Cells(ActiveCell.Row, startCol), Cells(ActiveCell.Row, endCol)).Interior.color = myInteriorColor
    End If
    
    '解放
    Set key = Nothing
    Set myRange = Nothing
    Set myNumF = Nothing
    
    Call 最適化もどす
    
    Call PlaySound("けってい")
    
End Function
Public Function QRコードをクリップボードに取得(Optional myString)
'    If IsMissing(myString) Then myString = "            0607         8211158560"
'    Dim MiBar As Mibarcd.Auto
'    Set MiBar = New Mibarcd.Auto
'    MiBar.CodeType = 12 '12=QR
'    MiBar.BarScale = 1
'    MiBar.QRVersion = 3 '大きくしたら大きくなる
'    MiBar.QRErrLevel = 1
'    MiBar.Code = myString
'    MiBar.Execute
End Function

Public Function フィールド名の追加(wsTemp, myKey, myArea, LR)
    retsu = myArea.count / 2
    With wsTemp
        For i = 1 To retsu
            myLR = myArea(i)
            If LR = "" Or myLR = "l" Then
                フィールド名 = myArea(retsu + i)
                Set mykey2 = .Cells.Find(フィールド名, , , 1)
                'フィールドが無い場合
                If mykey2 Is Nothing Then
                    .Columns(myKey.Column + i - 1).Insert
                    .Columns(myKey.Column + i - 1).Interior.Pattern = xlNone
                    .Cells(myKey.Row, myKey.Column + i - 1) = myArea(retsu + i)
                    .Columns(myKey.Column + i - 1).AutoFit
                    .Cells(myKey.Row, myKey.Column + i - 1).Interior.color = myArea(retsu + i).Interior.color
                    'コメントがある場合はコメント追加
                    If TypeName(myArea(retsu + i).Comment) <> "Nothing" Then
                        .Cells(myKey.Row, myKey.Column + i - 1).ClearComments
                        .Cells(myKey.Row, myKey.Column + i - 1).AddComment myArea(retsu + i).Comment.Text
                    End If
                'コメントがある場合はコメント削除してからコメント追加
                ElseIf TypeName(myArea(retsu + i).Comment) <> "Nothing" Then
                    .Cells(myKey.Row, myKey.Column + i - 1).ClearComments
                    .Cells(myKey.Row, myKey.Column + i - 1).AddComment myArea(retsu + i).Comment.Text
                End If
            End If
        Next i
    End With
End Function

Public Function 端末矢崎品番変換(端末矢崎品番)
    '-を含む場合は削除、含まない場合は-を付与
    If InStr(端末矢崎品番, "-") = 0 Then
        Select Case Len(端末矢崎品番)
        Case 8
            端末矢崎品番変換 = Left(端末矢崎品番, 4) & "-" & Mid(端末矢崎品番, 5, 4)
        Case 10
            端末矢崎品番変換 = Left(端末矢崎品番, 4) & "-" & Mid(端末矢崎品番, 5, 4) & "-" & Mid(端末矢崎品番, 9, 2)
        End Select
    Else
        端末矢崎品番変換 = Replace(端末矢崎品番, "-", "")
    End If
End Function

Public Function ポイントナンバー図作成(Optional 部品品番, Optional 端末, Optional 配列)
    Call 最適化
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim minW指定 As Long
    Dim myKey, actKey
    Dim cavCol As Long, ポイント1Col As Long, 二重係止Col As Long
    myFont = "ＭＳ ゴシック"
    minW指定 = 30
    'シートから呼び出した時
    If IsMissing(部品品番) Then
        With Workbooks(myBookName).Sheets(mySheetName)
            ハメ図タイプ = "チェッカー用"
            Set myKey = .Cells.Find("端末矢崎品番", , , 1)
            Set actKey = ActiveCell
            If actKey.Row <= myKey.Row Then Exit Function
            If .Cells(actKey.Row, myKey.Column) = "" Then Exit Function
            cavCol = .Cells.Find("Cav", , , 1).Column
            ポイント1Col = .Cells.Find("ポイント1", , , 1).Column
            二重係止Col = .Cells.Find("二重係止", , , 1).Column
            Dim 端末矢崎Col As Integer: 端末矢崎Col = .Cells.Find("端末矢崎品番", , , 1).Column
            Dim 端末Col As Integer: 端末Col = .Cells.Find("端末№", , , 1).Column
            Dim 略図col As Integer: 略図col = .Cells.Find("略図_表面視", , , 1).Column
            ReDim 配列(7, 0)
            部品品番 = .Cells(actKey.Row, 端末矢崎Col).Value
            端末 = .Cells(actKey.Row, 端末Col).Value
            Dim myCount1 As Long, myCount2 As Long
            Dim myTop As Long, myLeft As Long, myEnd As Long, myHeight As Long
            myCount1 = -1
            Do
                If 部品品番 <> .Cells(actKey.Row + myCount1, 端末矢崎Col) Or 端末 <> .Cells(actKey.Row + myCount1, 端末Col) Then
                    myTop = .Cells(actKey.Row + myCount1 + 1, 1).Top
                    myLeft = .Columns(略図col).Left
                    Exit Do
                End If
                myCount1 = myCount1 - 1
            Loop
            myCount2 = 1
            Do
                If 部品品番 <> .Cells(actKey.Row + myCount2, 端末矢崎Col) Or 端末 <> .Cells(actKey.Row + myCount2, 端末Col) Then
                    myEnd = .Cells(actKey.Row + myCount2, 1).Top
                    myHeight = myEnd - myTop
                    Exit Do
                End If
                myCount2 = myCount2 + 1
            Loop
            Dim y As Long, addc As Long
            For y = actKey.Row + myCount1 + 1 To actKey.Row + myCount2 - 1
                addc = UBound(配列, 2) + 1
                ReDim Preserve 配列(7, addc)
                配列(0, addc) = .Cells(y, cavCol)
                配列(1, addc) = .Cells(y, ポイント1Col)
                配列(2, addc) = .Cells(y, 二重係止Col)
            Next y
            部品品番 = 端末矢崎品番変換(部品品番)
            '画像がある場合は削除
            Dim objShp As Shape
            For Each objShp In ActiveSheet.Shapes
                If objShp.Name = 部品品番 & "_" & 端末 Then
                    objShp.Delete
                End If
            Next
        End With
    End If
    端末図 = 部品品番 & "_" & 端末
    Call addressSet(ActiveWorkbook)
    
    Dim 選択出力 As String
    Dim 倍率モード As Long: 倍率モード = 1 '0(現物倍) or 1(Cav基準倍)
    Dim 倍率 As Single
    Dim frameWidth As Long, frameWidth1 As Long, frameWidth2 As Long, frameHeight1 As Long, frameHeight2 As Long, cornerSize As Single
    Dim pp As Long

    Dim ハメ図種類 As String: ハメ図種類 = "写真" ' 写真(写真が無い時は略図) or 略図。拡張子はハメ図種類に応じて(固定)PVSW_RLTF両端にハメ図種類を出力する時に行う。
    Dim ハメ図拡張子 As String
    Dim ex As Long
    Dim varBinary As Variant
    Dim colHValue As New Collection  '連想配列、Collectionオブジェクトの作成
    Dim lngNu() As Long

    With Workbooks(myBookName).Sheets(mySheetName)
        
        '座標データの読込み(インポートファイル)
        Dim target As New FileSystemObject
        Dim targetDir As String: targetDir = myAddress(1, 1) & "\200_CAV座標"
        If Dir(targetDir, vbDirectory) = "" Then MsgBox "下記のファイルが無い為、各キャビティの座標が分かりません。" & vbCrLf & "部材一覧+で座標の出力を行ってから実行して下さい。" & vbCrLf & vbCrLf & myAddress(1, 1) & "\CAV座標.txt"
        
        Dim lastgyo As Long: lastgyo = 1
        Dim fileCount As Long: fileCount = 0
        Dim 使用部品str As String
        Dim 使用部品_端末 As String
        
        Dim aa As Variant, a As Variant
        Dim 座標発見Flag As Boolean
        Dim 使用部品_端末s_count As Long
        '使用部品Strに、今回使用する座標データを入れる
        Dim intFino As Variant
        intFino = FreeFile
        Dim 種類r(1) As String
        座標発見Flag = False
        種類r(0) = "png": 種類r(1) = "emf"
        minW = 1000: minH = 1000
        For ss = 0 To 1
            '写真,略図の順で探す
            URL = myAddress(1, 1) & "\200_CAV座標\" & 部品品番 & "_1_001_" & 種類r(ss) & ".txt"
            If Dir(URL) <> "" Then
                intFino = FreeFile
                Open URL For Input As #intFino
                Do Until EOF(intFino)
                    Line Input #intFino, aa
                    a = Split(aa, ",")
                    If a(0) <> "PartName" Then
                        For b = LBound(配列, 2) + 1 To UBound(配列, 2)
                            If CStr(配列(0, b)) = a(1) Then
                                配列(3, b) = a(2)
                                配列(4, b) = a(3)
                                配列(5, b) = a(4)
                                配列(6, b) = a(5)
                                配列(7, b) = a(7)
                                If minW > CLng(a(4)) Then minW = CLng(a(4))
                                If minH > CLng(a(5)) Then minH = CLng(a(5))
                                Exit For
                            End If
                        Next b
                    End If
                Loop
                Close #intFino
                Exit For
            End If
        Next ss
        Dim 使用部品 As Variant, 使用部品s As Variant, 使用部品c As Variant
line15:
        ReDim 電線データ(2, 1) As String
        '画像の配置
        ReDim 空栓表記(2, 0): 空栓c = 0
        Dim 画像無しflg As Boolean: 画像無しflg = False
        '写真
        画像URL = myAddress(1, 1) & "\201_写真\" & 部品品番 & "_1_" & Format(1, "000") & ".png"
        If Dir(画像URL) = "" Then
            '略図
            画像URL = myAddress(1, 1) & "\202_略図\" & 部品品番 & "_1_" & Format(1, "000") & ".emf"
            If Dir(画像URL) = "" Then
                画像無しflg = True 'GoTo line18
            End If
        End If
                                
        'If minW = -1 Then GoTo line18 'Cav座標が無ければ処理しない
        If 画像無しflg = True Then 'CAV座標にデータが無い時
            With ActiveSheet
                .Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 150, 60).Name = 端末図
                On Error Resume Next
                .Shapes.Range(端末図).Adjustments.Item(1) = 0.1
                On Error GoTo 0
                .Shapes.Range(端末図).Line.Weight = 1.6
                .Shapes.Range(端末図).TextFrame2.TextRange.Text = ""
                .Shapes.AddShape(msoShapeRoundedRectangle, 35, 10, 80, 40).Name = 端末図 & "_1"
                .Shapes.Range(端末図 & "_1").Adjustments.Item(1) = 0.1
                .Shapes.Range(端末図 & "_1").Line.Weight = 1.6
                .Shapes.Range(端末図 & "_1").TextFrame2.TextRange.Text = "no picture"
                .Shapes.Range(端末図).Select
                .Shapes.Range(端末図 & "_1").Select False
                Selection.Group.Select
                Selection.Name = 端末図
            End With
        ElseIf Dir(URL) = "" Then
            With ActiveSheet
                .Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 150, 60).Name = 端末図
                On Error Resume Next
                .Shapes.Range(端末図).Adjustments.Item(1) = 0.1
                On Error GoTo 0
                .Shapes.Range(端末図).Line.Weight = 1.6
                .Shapes.Range(端末図).TextFrame2.TextRange.Text = ""
                .Shapes.AddShape(msoShapeRoundedRectangle, 35, 10, 80, 40).Name = 端末図 & "_1"
                .Shapes.Range(端末図 & "_1").Adjustments.Item(1) = 0.1
                .Shapes.Range(端末図 & "_1").Line.Weight = 1.6
                .Shapes.Range(端末図 & "_1").TextFrame2.TextRange.Text = "座標.txtが無い"
                .Shapes.Range(端末図).Select
                .Shapes.Range(端末図 & "_1").Select False
                Selection.Group.Select
                Selection.Name = 端末図
            End With
        Else
            With ActiveSheet.Pictures.Insert(画像URL)
                .Name = 端末図
                If minW < minH Then
                    my幅 = (minW指定 / minW)
                Else
                    my幅 = (minW指定 / minH)
                End If
                .ShapeRange(端末図).ScaleHeight 1#, msoTrue, msoScaleFromTopLeft '画像が大きいとサイズを小さくされるから基のサイズに戻す
                .ShapeRange(端末図).ScaleHeight my幅, msoTrue, msoScaleFromTopLeft
                .CopyPicture
                .Delete
            End With
            DoEvents
            Sleep 10
            DoEvents
            .Paste
            Selection.Name = 端末図
            
            .Shapes(端末図).Left = 0
            .Shapes(端末図).Top = 0
            For i = LBound(配列, 2) + 1 To UBound(配列, 2)
                cav = 配列(0, i)
                If 配列(7, i) = "Ter" Then 配列(7, i) = "Box"
                If 配列(2, i) = True Or 配列(2, i) = 1 Then 二重係止flg = True Else 二重係止flg = False
                Call ColorMark3(端末, CStr(配列(3, i)), CStr(配列(4, i)), CStr(配列(5, i)), CStr(配列(6, i)), "", "", 配列(7, i), "", "", 配列(1, i) & vbLf & 配列(0, i), "", "", "", "", RowStr)
            Next i
            .Shapes.Range(端末図).Select
            For i = LBound(配列, 2) + 1 To UBound(配列, 2)
                .Shapes.Range(端末図 & "_" & 配列(0, i)).Select False
            Next i
            Selection.Group.Select
            Selection.Name = 端末図
            Selection.ShapeRange.Flip msoFlipHorizontal
            Selection.Copy
            Selection.Delete
            ActiveSheet.PasteSpecial Format:="図 (PNG)", Link:=False, DisplayAsIcon:=False
            Selection.Name = 端末図
        End If
        'シートから実行した時の処理
        If myTop <> 0 Then
            Selection.Left = myLeft
            Selection.Top = myTop
            Selection.Height = myHeight
            actKey.Select
        Else
            Set ポイントナンバー図作成 = Selection
        End If
    End With
    Call 最適化もどす
    
End Function
Public Function 後ハメ図呼び出し用QR印刷データ作成(Optional 治具str)
    If IsMissing(治具str) Then
        治具str = "152"
    End If
    Set wb(0) = ActiveWorkbook
    Set ws(0) = wb(0).Worksheets("冶具_" & 治具str)
    'ワークブック作成
    myBookpath = wb(0).path
    '出力先ディレクトリが無ければ作成
    If Dir(myBookpath & "\56_配索図_誘導", vbDirectory) = "" Then
        MkDir myBookpath & "\56_配索図_誘導"
    End If
    
    With ws(0)
        Set myKey = .Cells.Find("Size_", , , 1)
        Dim 端末Ran As Variant
        ReDim 端末Ran(0)
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For y = myKey.Row + 1 To lastRow
            xx = 1
            Do Until .Cells(y, xx) = ""
                セルstr = .Cells(y, xx).Value
                If Left(セルstr, 1) <> "U" Then
                    ReDim Preserve 端末Ran(UBound(端末Ran) + 1)
                    端末Ran(UBound(端末Ran)) = セルstr
                End If
                xx = xx + 2
            Loop
        Next y
    End With
    
    newBookName = "QR印刷_" & 治具str & ".xlsx"
    Set wb(1) = Workbooks.add
    
    With wb(1).Sheets("Sheet1")
        .Cells.NumberFormat = "@"
        .Cells(1, 1) = "QR"
        .Cells(1, 2) = "端末"
        .Cells(2, 2) = "治具_" & 治具str
        addRow = 3
        For y = LBound(端末Ran) + 1 To UBound(端末Ran)
            .Cells(addRow, 1).Resize(1, 2) = 端末Ran(y)
            addRow = addRow + 1
        Next y
    End With
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=myBookpath & "\56_配索図_誘導\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    wb(1).Close
    
End Function

Public Function 誘導モニタの移動データ作成_後ハメ図csv(製品品番str, 手配str, 冶具str)
    'temp
    Set myBook = ActiveWorkbook
    Dim 端末一覧ran()
    Call SQL_端末一覧(端末一覧ran, 製品品番str, myBook.Name)

    With myBook.Sheets("冶具_" & 冶具str)
        .Activate
        Dim moveX As Long, moveXpt As Single
        Dim 冶具Wmm As Single: 冶具Wmm = .Cells.Find("Width_", , , 1).Offset(0, 1)
        If 冶具Wmm = 0 Then Stop '治具Wmmが入力されていません
        Dim 冶具Wpt As Single: 冶具Wpt = .Shapes.Range("板").Width
        Dim 後ハメ図dir As String
        後ハメ図dir = ActiveWorkbook.path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\xMove"
        If Dir(後ハメ図dir, vbDirectory) = "" Then MkDir (後ハメ図dir)
        Dim 後ハメ図path As String
        後ハメ図path = 後ハメ図dir & "\後ハメ図.csv"
        Open 後ハメ図path For Output As #1
        For i = LBound(端末一覧ran, 2) To UBound(端末一覧ran, 2)
            端末str = 端末一覧ran(1, i)
            moveXpt = .Shapes.Range(端末str).Left + (.Shapes.Range(端末str).Width / 2)
            moveX = moveXpt / 冶具Wpt * 冶具Wmm
            Print #1, 端末str & "," & moveX & "," & 端末一覧ran(3, i)
        Next i
        Close #1
    End With
    Dim array_temp(0, 0) As Variant
    array_temp(0, 0) = 冶具Wmm
    export_Array_ShiftJis_ver2 array_temp, 後ハメ図dir & "\width.txt", ","
End Function
Public Function 誘導モニタの移動データ作成_構成_構成の中心csv(製品品番str, 手配str, 冶具str)
    'temp
    Set myBook = ActiveWorkbook
    Call SQL_配策図用_回路(配索端末RAN, 製品品番str, myBook)
    'Call SQL_端末一覧(端末一覧ran, 製品品番str, myBook.Name)
    Dim サブ座標RAN()
    ReDim サブ座標RAN(2, 0)

    Dim myArrry As Variant
    myArrry = readSheetToRan3(wb(0).Sheets("PVSW_RLTF"), "電線識別名", "構成,始点側端末,終点側端末,始点側ハメ,終点側ハメ", "", 1)
    
    With myBook.Sheets("冶具_" & 冶具str)
        Dim moveX As Long, moveXpt As Single
        Dim 冶具Wmm As Single: 冶具Wmm = .Cells.Find("Width_", , , 1).Offset(0, 1)
        Dim 冶具Wpt As Single: 冶具Wpt = .Shapes.Range("板").Width
        'サブ毎の中心ptを求めてサブ座標ranに格納
        Dim minX As Single, maxX As Single, aveX As Single, Status As String, Result As String, Resultpt As Single
        For i = LBound(配索端末RAN, 2) + 1 To UBound(配索端末RAN, 2)
            Aハメ = 配索端末RAN(7, i)
            Bハメ = 配索端末RAN(9, i)
            If IsNull(Aハメ) Then Aハメ = ""
            If IsNull(Bハメ) Then Bハメ = ""
            Status = Aハメ & Bハメ
            If Status = "後後" Then
                Result = "ave"
            ElseIf Status Like "後*" Then
                Result = 配索端末RAN(4, i)
            ElseIf Status Like "*後" Then
                Result = 配索端末RAN(5, i)
            ElseIf Status = "先ハメ先ハメ" Then
                Result = "ave"
            ElseIf Status Like "先ハメ*" Then
                Result = 配索端末RAN(4, i)
            ElseIf Status Like "*先ハメ" Then
                Result = 配索端末RAN(5, i)
            Else
                Result = "ave"
            End If
            For x = 0 To 1 '始点終点の端末
                端末str = 配索端末RAN(4 + x, i)
                If 端末str <> "" Then
                    On Error Resume Next
                    端末pt = .Shapes.Range(端末str).Left + (.Shapes.Range(端末str).Width / 2)
                    If err.number = 1004 Then Exit Function
                    On Error GoTo 0
                    If 端末pt < minX Or minX = 0 Then minX = 端末pt
                    If 端末pt > maxX Then maxX = 端末pt
                    If 端末str = Result Then Resultpt = 端末pt
                End If
                If x = 1 Then
                    If minX = 0 Then minX = maxX
                    If maxX = 0 Then maxX = minX
                    If Result = "ave" Then
                        aveX = minX + ((maxX - minX) / 2)
                    Else
                        aveX = Resultpt
                    End If
                    ReDim Preserve サブ座標RAN(2, UBound(サブ座標RAN, 2) + 1)
                    サブ座標RAN(0, UBound(サブ座標RAN, 2)) = aveX
                    サブ座標RAN(1, UBound(サブ座標RAN, 2)) = minX
                    サブ座標RAN(2, UBound(サブ座標RAN, 2)) = maxX
                    minX = 0: maxX = 0
                End If
            Next x
        Next i
        
        Dim 後ハメ図dir As String
        後ハメ図dir = ActiveWorkbook.path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\xMove"
        If Dir(後ハメ図dir, vbDirectory) = "" Then MkDir (後ハメ図dir)
        Dim 後ハメ図path As String
        後ハメ図path = 後ハメ図dir & "\構成.csv"
        Open 後ハメ図path For Output As #1
        For i = LBound(配索端末RAN, 2) + 1 To UBound(配索端末RAN, 2)
            構成str = 配索端末RAN(2, i)
            'サブstr = 配索端末RAN(1, i)
            aveX = サブ座標RAN(0, i) / 冶具Wpt * 冶具Wmm
            minX = サブ座標RAN(1, i) / 冶具Wpt * 冶具Wmm
            maxX = サブ座標RAN(2, i) / 冶具Wpt * 冶具Wmm
            colorLong = 配索端末RAN(11, i)
            Print #1, 構成str & "," & aveX & "," & minX & "," & maxX & "," & colorLong
        Next i
        Close #1
    End With
End Function
Public Function 誘導モニタの移動データ作成_構成_サブの中心csv(製品品番str, 手配str, 冶具str)
    'temp
    Set myBook = ActiveWorkbook
    Call SQL_配策図用_回路(配索端末RAN, 製品品番str, myBook)
    'Call SQL_端末一覧(端末一覧ran, 製品品番str, myBook.Name)
    Dim サブ座標RAN()
    ReDim サブ座標RAN(1, 0)
    
    With myBook.Sheets("冶具_" & 冶具str)
        Dim moveX As Long, moveXpt As Single
        Dim 冶具Wmm As Single: 冶具Wmm = .Cells.Find("Width_", , , 1).Offset(0, 1)
        Dim 冶具Wpt As Single: 冶具Wpt = .Shapes.Range("板").Width
        
        'サブ毎の中心ptを求めてサブ座標ranに格納
        Dim minX As Single, maxX As Single, aveX As Single, insertStatus As String
        For i = LBound(配索端末RAN, 2) + 1 To UBound(配索端末RAN, 2)
            サブstr = 配索端末RAN(1, i)

            
            For x = 0 To 1 '始点終点の端末
                端末str = 配索端末RAN(4 + x, i)
            
                If 端末str <> "" Then
                    On Error Resume Next
                    端末pt = .Shapes.Range(端末str).Left + (.Shapes.Range(端末str).Width / 2)
                    If err.number = 1004 Then Exit Function
                    On Error GoTo 0
                    If 端末pt < minX Or minX = 0 Then minX = 端末pt
                    If 端末pt > maxX Then maxX = 端末pt
                End If
                
                If x = 1 Then
                    If i = UBound(配索端末RAN, 2) Then
                        aveX = minX + ((maxX - minX) / 2)
                        ReDim Preserve サブ座標RAN(1, UBound(サブ座標RAN, 2) + 1)
                        サブ座標RAN(0, UBound(サブ座標RAN, 2)) = サブstr
                        サブ座標RAN(1, UBound(サブ座標RAN, 2)) = aveX
                        minX = 0: maxX = 0
                    Else
                        サブnext = 配索端末RAN(1, i + 1)
                        If サブstr <> サブnext Then
                            aveX = minX + ((maxX - minX) / 2)
                            ReDim Preserve サブ座標RAN(1, UBound(サブ座標RAN, 2) + 1)
                            サブ座標RAN(0, UBound(サブ座標RAN, 2)) = サブstr
                            サブ座標RAN(1, UBound(サブ座標RAN, 2)) = aveX
                            minX = 0: maxX = 0
                        End If
                    End If
                End If
            Next x
        Next i
        
        Dim 後ハメ図dir As String
        後ハメ図dir = ActiveWorkbook.path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\xMove"
        If Dir(後ハメ図dir, vbDirectory) = "" Then MkDir (後ハメ図dir)
        Dim 後ハメ図path As String
        後ハメ図path = 後ハメ図dir & "\サブ.csv"
        Open 後ハメ図path For Output As #1
        For i = LBound(配索端末RAN, 2) + 1 To UBound(配索端末RAN, 2)
            構成str = 配索端末RAN(2, i)
            サブstr = 配索端末RAN(1, i)
            For ii = LBound(サブ座標RAN, 2) + 1 To UBound(サブ座標RAN, 2)
                If サブstr = サブ座標RAN(0, ii) Then
                    moveXpt = サブ座標RAN(1, ii)
                    moveX = moveXpt / 冶具Wpt * 冶具Wmm
                    Print #1, 構成str & "," & moveX & "," & サブstr
                    Exit For
                End If
            Next ii
        Next i
        Close #1
    End With
End Function


Sub MakeShortcut(path)
    Dim ShellObject
    Set ShellObject = CreateObject("WScript.Shell")
   
    Dim ShortcutObject
    Set ShortcutObject = ShellObject.CreateShortcut(path & "\" & ActiveWorkbook.Name & ".lnk")
    
    With ShortcutObject
        .TargetPath = ActiveWorkbook.FullName
        .Save
    End With
End Sub

Sub ログ出力test_temp()
    Call addressSet(wb(0))
    Set wb(0) = ThisWorkbook
    Call ログ出力("aaa", "bbb", "textttttttttttt")
End Sub

Public Function ログ出力(フォルダ, ファイル名, テキスト1)
    Dim myPath As String, myIP As String, myDir As String
    myPath = myAddress(0, 1) & "\log\" & フォルダ & "\" & ファイル名 & ".txt"
    myDir = myAddress(0, 1) & "\log\" & フォルダ
    myIP = GetIPAddress
    'フォルダが無ければ作成
    If Dir(myDir, vbDirectory) = "" Then
        MkDir (myDir)
    End If
    'テキストファイルを使用する
    Dim tFso As FileSystemObject
    Dim tFile As TextStream
    'ファイルが無ければ新規作成
    If Dir(myPath) = "" Then
        Set tFso = New FileSystemObject
        Set tFile = tFso.CreateTextFile(myPath, True)
        Set tFso = Nothing
        Set tFile = Nothing
    End If
                
    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをAppendモードで開きます。
    Open myPath For Append As #FileNumber
    Dim outdats(3)
    '出力用の配列へデータをセットします。
    outdats(0) = now
    outdats(1) = myIP
    outdats(2) = テキスト1
    outdats(3) = ThisWorkbook.FullName
    '配列の要素をカンマで結合して出力します。
    Print #FileNumber, Join(outdats, vbTab)

    '入力ファイルを閉じます。
    Close #FileNumber
    
End Function

Public Function 部材詳細の読み込み(部品品番str, フィールド名str)
        Dim target As New FileSystemObject
        Dim path As String: path = myAddress(1, 1) & "\300_部材詳細\" & 部品品番str & ".txt"
        If Dir(path) = "" Then 部材詳細の読み込み = False: Exit Function
        Dim intFino As Variant
        intFino = FreeFile
        Open path For Input As #intFino
        myX = ""
        Do Until EOF(intFino)
            Line Input #intFino, aa
            temp = Split(aa, ",")
            For x = LBound(temp) To UBound(temp)
                If Replace(temp(x), "-", "") = フィールド名str Then
                    Line Input #intFino, aa
                    temp = Split(aa, ",")
                    部材詳細の読み込み = temp(x)
                    Close #intFino
                    Exit Function
                End If
            Next x
        Loop
        Close #intFino
End Function

Public Function 部材詳細の読み込みVer2(部品品番str, フィールド名str) As String
        Dim target As New FileSystemObject
        Dim path As String: path = myAddress(1, 1) & "\300_部材詳細\" & 部品品番str & ".txt"
        If Dir(path) = "" Then 部材詳細の読み込みVer2 = False: Exit Function
        Dim sp As Variant, x As Long, x2 As Long
        sp = Split(フィールド名str, ",")
        
        Dim intFino As Variant, temp As Variant, temp2 As Variant, tempStr As String, aa As String, aa2 As String, temp3 As Variant
        intFino = FreeFile
        Open path For Input As #intFino
        Do Until EOF(intFino)
            Line Input #intFino, aa
            temp = Split(aa, ",")
            Line Input #intFino, aa
            temp2 = Split(aa, ",")
            ReDim temp3(UBound(sp))
            
                For x2 = LBound(sp) To UBound(sp)
            For x = LBound(temp) To UBound(temp)
                    If Replace(temp(x), "-", "") = sp(x2) Then
                        temp3(x2) = temp2(x)
                        Exit For
                    End If
            Next x
                Next x2
            
        Loop
        Close #intFino
        tempStr = Join(temp3, ",")
        部材詳細の読み込みVer2 = tempStr
End Function

Public Function searchCAV座標(部品品番str, フィールド名str)
        Dim target As New FileSystemObject
        Dim path As String: path = myAddress(1, 1) & "\200_CAV座標\" & 部品品番str & "_1_001_png.txt"
        If Dir(path) = "" Then path = myAddress(1, 1) & "\200_CAV座標\" & 部品品番str & "_1_001_emf.txt"
        If Dir(path) = "" Then searchCAV座標 = False: Exit Function
        Dim intFino As Variant
        intFino = FreeFile
        Open path For Input As #intFino
        myX = ""
        Do Until EOF(intFino)
            Line Input #intFino, aa
            temp = Split(aa, ",")
            For x = LBound(temp) To UBound(temp)
                If Replace(temp(x), "-", "") = フィールド名str Then
                    Line Input #intFino, aa
                    temp = Split(aa, ",")
                    searchCAV座標 = temp(x)
                    Close #intFino
                    Exit Function
                End If
            Next x
        Loop
        Close #intFino
End Function

Public Function セルの中身を全て渡す(Base As Range, aite As Range)
    Base.Value = aite.Value
    If aite.Interior.ColorIndex <> xlNone Then Base.Interior.color = aite.Interior.color
    If Not (aite.Comment Is Nothing) Then
        Set コメント = Base.AddComment
        コメント.Text aite.Comment.Text
        コメント.Visible = False
        コメント.Shape.Fill.ForeColor.RGB = RGB(255, 192, 0)
        コメント.Shape.TextFrame.AutoSize = True
        コメント.Shape.TextFrame.Characters.Font.size = 11
        コメント.Shape.Placement = xlMove
    End If
End Function

Public Function TEXT出力_配索経路_端末経路css(myPath)
    
    Dim FileNumber      As Integer
    Dim outdats(1 To 47) As Variant
    Dim myRow           As Long
    Dim i               As Integer
    Dim flg_out         As Boolean

    'box2l = 1.1218 * ((box2l * 100) ^ 0.9695)
    'box2l = (0.9898 * (box2l * 100)) + 0.2766
    'box2t = 1.0238 * ((box2t * 100) ^ 0.9912)

    '空いているファイル番号を取得します。
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open myPath For Output As #FileNumber
    
        outdats(1) = "table {"
        outdats(2) = "    table-layout: fixed;"
        outdats(3) = "    width:100%;"
        outdats(4) = "    background-color:#232526;"
        outdats(5) = "    text-align:center;"
        outdats(6) = "    color: #FFFFFF;"
        outdats(7) = "    font-size:14pt;"
        outdats(8) = "    font-weight: bold;"
        outdats(9) = "    border-collapse: collapse;"
        outdats(10) = "    font-family: Verdana, Arial, Helvetica, sans-serif;"
        outdats(11) = "    border-bottom:0px solid #000000;"
        outdats(12) = "}"
        outdats(13) = "table td {"
        outdats(14) = "    border: 1px solid  #FFFFFF;"
        outdats(15) = "    border-left:2px solid #FFFFFF;"
        outdats(16) = "    border-right:2px solid  #FFFFFF;"
        outdats(17) = "    padding: 1px;"
        outdats(18) = "}"
        outdats(19) = ".box1 img{"
        outdats(21) = "    position:absolute;"
        outdats(22) = "    width:99%;"
        outdats(23) = "    height:auto;"
        outdats(24) = "    max-width:99%;"
        outdats(25) = "    max-height:95%;"
        outdats(26) = "}"
        outdats(27) = ".box1 {"
        outdats(28) = "}"
        outdats(29) = "#box2 img{"
        outdats(30) = "    filter:alpha(opacity=70);"
        outdats(31) = "    position: absolute;"
        outdats(32) = "    width:99%;"
        outdats(33) = "    opacity:0.7;"
        outdats(34) = "    zoom:1;"
        outdats(35) = "    display:inline-block;"
        outdats(36) = "}"
        outdats(37) = "#box3 img{"
        outdats(38) = "    position:absolute;"
        outdats(39) = "    width:99%;"
        outdats(40) = "}"
        outdats(41) = "#box4 img{"
        outdats(42) = "    position:absolute;"
        outdats(43) = "    bottom:0%;"
        outdats(44) = "    width:99%;"
        outdats(45) = "}"
        outdats(46) = "body{background-color:#111111;}"
        outdats(47) = ".myB{color:#FFFFFF;background-color:#232526;}"
        
        '配列の要素をカンマで結合して出力します。
        Print #FileNumber, Join(outdats, vbCrLf)

    '入力ファイルを閉じます。
    Close #FileNumber

End Function

Public Sub fieldAdd(ByVal sheetName As String, ByVal fieldString As String, ByVal myInt As Integer)
    '対象シートが無ければ抜ける
    If checkSheet(sheetName, wb(0), False, False) = False Then Exit Sub
    
    With ActiveWorkbook.Sheets("フィールド名")
        Set myKey = .Cells.Find(fieldString, , , 1)
        Dim lastCol As Long
        lastCol = .Cells(myKey.Row + 2, .Columns.count).End(xlToLeft).Column
        Set myArea = .Range(myKey.Offset(1, 0).addRess, myKey.Offset(2, lastCol).addRess)
    End With
    
    With ActiveWorkbook.Sheets(sheetName)
        Set myKey = .Cells.Find(myArea(myInt, 1), , , 1)
        If myKey Is Nothing Then
            MsgBox ("[製品品番] にフィールド: " & myArea(myInt, 1) & "が見つかりません｡" & vbLf & "処理を中断します")
            End
        End If
        retsu = myArea.count / myInt
        For i = 1 To retsu
            フィールド名 = myArea(myInt, i)
            Set mykey2 = .Cells.Find(フィールド名, , , 1)
            'フィールドが無い場合
            If mykey2 Is Nothing Then
                .Columns(myKey.Column + i - 1).Insert
                .Columns(myKey.Column + i - 1).Interior.Pattern = xlNone
                .Cells(myKey.Row, myKey.Column + i - 1) = myArea(myInt, i)
                If myInt > 1 Then .Cells(myKey.Row - 1, myKey.Column + i - 1) = myArea(i)
                If myArea(myInt, i).Interior.Pattern = 1 Then
                    .Cells(myKey.Row, myKey.Column + i - 1).Interior.color = myArea(myInt, i).Interior.color
                End If
                'コメントがある場合はコメント追加
                If TypeName(myArea(retsu + i).Comment) <> "Nothing" Then
                    .Cells(myKey.Row, myKey.Column + i - 1).ClearComments
                    .Cells(myKey.Row, myKey.Column + i - 1).AddComment myArea(retsu + i).Comment.Text
                End If
            'コメントがある場合はコメント削除してからコメント追加
            ElseIf TypeName(myArea(retsu + i).Comment) <> "Nothing" Then
                .Cells(myKey.Row, myKey.Column + i - 1).ClearComments
                .Cells(myKey.Row, myKey.Column + i - 1).AddComment myArea(retsu + i).Comment.Text
            End If
        Next i
    End With
End Sub

Public Function 座標状況更新()
    Call 最適化
    
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "04_部材一覧"

    Dim 座標path As String: 座標path = myAddress(0, 1) & "\200_CAV座標"

    With myBook.Sheets(mySheetName)
        Dim inKey As Range: Set inKey = .Cells.Find("部品品番_")
        Dim startRow As Long: startRow = inKey.Row + 1
        Dim my部品品番Col As Long: my部品品番Col = inKey.Column
        Dim my区分Col As Long: my区分Col = .Cells.Find("区分_").Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, my部品品番Col).End(xlUp).Row
        Set inKey = Nothing
        'Dim 登録座標RAN
        'Call SQL_登録済み座標の取得(登録座標RAN, myBook)
        Dim i, found As Long
        Dim 種類 As String, 部品品番str As String
        For i = startRow To lastRow
            If .Cells(i, my区分Col) = 4 Or .Cells(i, my区分Col) = 3 Then
                部品品番str = .Cells(i, my部品品番Col)
                For C = 1 To 2
                    If C = 1 Then 種類 = "emf" Else 種類 = "png"
                    If Dir(座標path & "\" & 部品品番str & "_1_001_" & 種類 & ".txt") <> "" Then
                        found = "1"
                    Else
                        found = "0"
                    End If
                    .Cells(i, C) = found
                Next C
            End If
        Next i
    End With
    Call 最適化もどす
    MsgBox "座標状況を更新しました" & vbLf & "経過時間:= " & DateDiff("s", jj, time)
End Function

Function search親端末(ByVal 端末1 As String, ByVal 端末2 As String, ByVal 接続G As String, ByVal myRan As Variant, ByVal d As Variant, ByVal dd As Variant) As String

    Dim i As Long, foundFlag As Boolean, swapStr As String, swapFlag As Boolean
    Dim p(3) As Variant
    For i = 0 To UBound(d)
        If dd(i) = "親端末No" Then p(0) = i
        If dd(i) = "始点側端末識別子" Then p(1) = i
        If dd(i) = "終点側端末識別子" Then p(2) = i
        If dd(i) = "接続G_" Then p(3) = i
    Next i
    
line10:
    For i = LBound(myRan, 2) To UBound(myRan, 2)
        If 端末1 = myRan(p(1), i) And 端末2 = myRan(p(2), i) And 接続G = myRan(p(3), i) Then
            search親端末 = myRan(p(0), i)
            Exit Function
        End If
    Next i
    
    '見つからない場合は条件を入れ替える
    If swapFlag = False Then
        swapStr = 端末2
        端末2 = 端末1
        端末1 = swapStr
        swapFlag = True
        GoTo line10
    End If
    
    'それでも見つからない場合
    search親端末 = False
    Stop 'この時の処理要
End Function

Sub UI_show_search()
    UI_search.Show
End Sub


Function readText_UTF8(ByVal myPath As String)
    Dim target As String, buf As String
    target = myPath
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile target
        buf = .ReadText
        .Close
        readText_UTF8 = buf
    End With
End Function

Public Function exportHTML_navi_UTF8(ByVal mytext As Variant, ByVal myDir As String)
    
    Dim txtFile As String, i As Long
    Dim adoSt As ADODB.Stream
    Set adoSt = New ADODB.Stream
    Dim strLine As String, mytextSP As Variant
    mytextSP = Split(mytext, vbLf)
    
    With adoSt
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        For i = LBound(mytextSP) To UBound(mytextSP)
            strLine = mytextSP(i)
            .WriteText strLine, adWriteLine
        Next i
        'ここからBOM無しにする処理
        .position = 0
        .Type = adTypeBinary
        .position = 3 'BOMデータは3バイト目まで
        Dim byteData() As Byte '一時格納
        byteData = .Read  '一時格納用変数に保存
        .Close 'ストリームを閉じる_リセット
        .Open
        .Write byteData
        .SaveToFile myDir, adSaveCreateOverWrite
        .Close
    End With
        
End Function
