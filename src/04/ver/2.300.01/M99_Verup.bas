Attribute VB_Name = "M99_Verup"
Function VBC_Export(Optional path)
    Dim VBC
    Dim myCount As Long

    With ActiveWorkbook.VBProject
        For Each VBC In .VBComponents
            Debug.Print VBC.Type, VBC.Name
            If VBC.Type <> 100 And _
               VBC.CodeModule.CountOfDeclarationLines <> VBC.CodeModule.CountOfLines Then
                If VBC.Type = 1 Then VBC.Export path & "\" & VBC.Name & ".bas"
                If VBC.Type = 2 Then VBC.Export path & "\" & VBC.Name & ".cls"
                If VBC.Type = 3 Then VBC.Export path & "\" & VBC.Name & ".frm"
                myCount = myCount + 1
            End If
        Next VBC
    End With
    VBC_Export = myCount
End Function
'
'Sub VBC_Remove() 'このモジュールは削除しない
'    Dim VBC
'    With ActiveWorkbook.VBProject
'        For Each VBC In .VBComponents
'            Debug.Print VBC.Type, VBC.Name
'            If VBC.Type <> 100 Then
'                .VBComponents.Remove .VBComponents(VBC.Name)
'            End If
'        Next VBC
'    End With
'End Sub
'Public Function VBC_Import()
'    Path = "D:\04_生産準備+\000_システムパーツ\ver\2.125"
'    Dim buf As String, cnt As Long
'    Path = Path & "\"
'    buf = Dir(Path & "*.*")
'    Do While buf <> ""
'        'Debug.Print buf
'        If Right(buf, 3) <> "frx" And Right(buf, 3) <> "log" Then
'            ActiveWorkbook.VBProject.VBComponents.Import (Path & buf)
'        End If
'        buf = Dir()
'    Loop
'End Function

Sub Sheet_Ver_Export(path)
    DeleteDefinedNames
    
    ThisWorkbook.Sheets("Ver").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=path & "\sheet_Ver", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
    
    ThisWorkbook.Sheets("フィールド名").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=path & "\sheet_FieldName", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
    
    ThisWorkbook.Sheets("color").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=path & "\sheet_color", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
    
    ThisWorkbook.Sheets("設定").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=path & "\sheet_setting", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
    
    ThisWorkbook.Sheets("WEB").Visible = True
    ThisWorkbook.Sheets("WEB").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=path & "\sheet_web", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
    
    ThisWorkbook.Sheets("効果").Visible = True
    ThisWorkbook.Sheets("効果").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=path & "\sheet_effect", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
End Sub
