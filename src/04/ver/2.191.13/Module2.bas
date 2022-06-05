Attribute VB_Name = "Module2"
 Sub CommandButton������2��()
    ActiveSheet.Range("D2").Select
'�i*���̕����ȗ�*�L�^�}�N������ύX2�@�̃R�[�h�����̂܂ܓ\��t����j
    ActiveCell.Offset(0, 3).Activate
    With ActiveSheet.OLEObjects.add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False)
        .Object.Caption = "menu"
        .Object.Font.Name = "���C���I"
        .Object.Font.Size = 13
        .Width = ActiveCell.Width * 2
        .Height = ActiveCell.Height  'ActiveCell.Height * 2
    End With
End Sub
Sub SetButtonsOnActiveSheet(codeName As String)
    'For Excel VBA
    'Buid by Q11Q From Qiita
    'https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q11106096531
    '[�}�N���Ń{�^����}������ - moug](https://www.moug.net/tech/exvba/0150104.html)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim shp As Excel.Shape
    'Dim r As Range: Set r = ws.Range("G9:E9") '�A�������Z���ł����������Z���ł��悢
    Dim r As Range: Set r = ws.Range("c1")
    Dim obj As Object
    Dim shpRng As Excel.ShapeRange
    With r
        Set obj = ws.Buttons.add(.Left, .Top, Application.CentimetersToPoints(3.65), Application.CentimetersToPoints(0.56)) '�v���p�e�B�V�[�g��Cm�P�ʂȂ̂ŁACm�P�ʂŎw�肵�Ă���B���������ۂ͌덷���o�邽�߁A�o�͂��ꂽ�{�^���̃T�C�Y�͓����ɂȂ�Ȃ��B
        obj.Caption = "MENU" 'Caption�ł��w��ł���
        obj.OnAction = codeName '�}�N���̖��O������
        'obj.Enable = True '�g�p�\�ɂ��� �������G���[�ɂȂ邽�ߎg���Ȃ�
        obj.PrintObject = False 'False ������Ȃ�
        obj.Visible = True '��
        obj.Font.Name = "���C���I"
        obj.Font.FontStyle = "�W��"
        obj.Font.Size = 12
        obj.Font.Underline = xlUnderlineStyleNone
        obj.Locked = True
        obj.LockedText = True
        obj.Width = .Width
        obj.Height = .Height * 2
    End With
    Set shp = ws.Shapes(ws.Shapes.count)
    shp.ZOrder msoBringToFront '�őO�ʂɈړ�
    Set shpRng = ws.Shapes.Range(Array(shp.Name))
    'Stop
End Sub
Sub InsertButtonOnSheet()
' For Excel VBA
' �����L�^���Q�l�ɍ쐬�����o�[�W����
'F3�Ƀ{�^����ݒu����
'A1�ɍ����̓��t����͂���}�N�����N������悤�ɂ���
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet: Set ws = ActiveSheet
Dim Rng As Range
Dim obj As Object
Dim shp As Excel.Shape
Dim shpRng As Excel.ShapeRange
Set Rng = ws.Range("F3")
Rng.Select
Set obj = ws.Buttons.add(273.75, 40.5, 90.75, 15.75)
With obj
    obj.Characters.Text = "text editing"
    .OnAction = "sampleInsertToday"
    With obj.Characters(Start:=1, Length:=12).Font
        .Name = "�l�r �o�S�V�b�N"
        .FontStyle = "�W��"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
End With
Set shp = ws.Shapes(ws.Shapes.count)
shp.ZOrder msoBringToFront '�őO�ʂɈړ�
Set shpRng = ws.Shapes.Range(Array(shp.Name))
shpRng.Select
    With Selection.Font
        .Name = "�l�r �o�S�V�b�N"
        .FontStyle = "�W��"
        .Size = 10
        .Strikethrough = True
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .Orientation = xlHorizontal
        .AddIndent = False
    End With
    With obj
        .Placement = xlFreeFloating
        .PrintObject = True 'object����������
    End With
End Sub


Sub sampleInsertToday()
'A1�ɍ����̓��t����͂���}�N��
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet: Set ws = ActiveSheet
Dim Rng As Range: Set Rng = ws.Range("A1"): Rng.Activate
Rng.Value = Date
End Sub

