Attribute VB_Name = "M11_�`��"
'�X���[�v�C�x���g
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' �萔�̐錾

Public GYO As Long
Public retsu As Long

Public ColorVal() As String     ' �J���[�l�i�[�p
Public ColorValFont() As Long ' �t�H���g�p
Public ColorCode() As Long    ' �J���[�l�i�[�p
Public ColorName() As String  ' �F�L���i�[�p

Public WhiteLineFrg As Boolean      ' �����t���O�i�h��F�����̏ꍇ���̐F�𔒂ɂ���j
Public �[�� As String
Public �[���} As String
Public cav As Long
Public �[��cav�W�� As String
Public my�� As Single


Public Function ColorNameToColorValue(color, filcolor, strcolor, fontcolor)
    
        If InStr(color, "/") > 0 Then
            filc = Left(color, InStr(color, "/") - 1)
            strc = Mid(color, InStr(color, "/") + 1)
        Else
            filc = color
            strc = filc
        End If
        
        colorFind = False
        For qq = LBound(ColorCode) To UBound(ColorCode)
            If ColorName(qq) = filc Then
                filcolor = ColorVal(qq)
                fontcolor = ColorValFont(qq)
                colorFind = True
                Exit For
            End If
        Next qq
        If colorFind = False Then Stop '�F�̓o�^������
        
        If filc <> strc Then
            colorFind = False
            For qq = LBound(ColorCode) To UBound(ColorCode)
                If ColorName(qq) = strc Then
                    strcolor = ColorVal(qq)
                    colorFind = True
                    Exit For
                End If
            Next qq
            If colorFind = False Then Stop '�F�̓o�^������
        Else
            strcolor = filcolor
        End If
End Function

Sub Init2()
    With myBook.Sheets("color")
        Set key = .Cells.Find("ColorName", , , 1)
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        
        ReDim ColorName(lastRow - key.Row)
        ReDim ColorCode(lastRow - key.Row)
        ReDim ColorVal(lastRow - key.Row)
        ReDim ColorValFont(lastRow - key.Row)
        
        For i = key.Row + 1 To lastRow
            ColorName(i - key.Row - 1) = .Cells(i, 1)
            ColorCode(i - key.Row - 1) = .Cells(i, 2)
            ColorVal(i - key.Row - 1) = .Cells(i, 3)
            ColorValFont(i - key.Row - 1) = .Cells(i, 4)
        Next i
    End With

End Sub

Sub Init()
    
    ' *** ����������
    ' �J���[�l�i�o�b��ł̕\���l�j
    ColorVal(0) = -1                   ' ����
    ColorVal(1) = RGB(20, 20, 20)      ' ��
    ColorVal(2) = RGB(252, 252, 252)   ' ��
    ColorVal(3) = RGB(240, 0, 0)       ' ��
    ColorVal(4) = RGB(0, 186, 84)      ' ��
    ColorVal(5) = RGB(255, 255, 0)     ' ��
    ColorVal(6) = RGB(162, 89, 0)      ' ��
    ColorVal(7) = RGB(0, 110, 255)     ' ��
    ColorVal(8) = RGB(255, 160, 177)   ' �s���N
    ColorVal(9) = RGB(186, 186, 186)   ' �D
    ColorVal(10) = RGB(170, 255, 0)    ' ��t
    ColorVal(11) = RGB(101, 226, 255)  ' ��
    ColorVal(12) = RGB(186, 68, 255)   ' ��
    ColorVal(13) = RGB(255, 130, 17)    ' �I�����W
    ColorVal(14) = RGB(205, 152, 0)    ' �`���R���[�g
    ColorVal(15) = RGB(255, 179, 102)  '�x�[�W��
    ColorVal(16) = RGB(100, 100, 100)  'ZZ(�F������o���Ȃ�)
    ColorVal(17) = RGB(93, 93, 93)     '�[�D
    ColorVal(18) = RGB(173, 173, 173)  '��D
    ColorVal(19) = RGB(203, 203, 203)  '��
    ColorVal(20) = RGB(6, 52, 6)       '�[��
    ColorVal(21) = RGB(255, 239, 143)  '�N���[��
    ColorVal(22) = RGB(234, 234, 89)   '����
    ColorVal(23) = RGB(6, 6, 74)       '�[��
    ColorVal(24) = RGB(63, 6, 0)       '�_�[�N�`���R
    ColorVal(25) = RGB(214, 182, 65)   '����
    ColorVal(26) = RGB(100, 74, 141)   '����
    ColorVal(27) = RGB(184, 101, 204)  '���x���_�[
    ColorVal(28) = RGB(230, 90, 0)     '���炵�F(����Ŏg�p)
    ColorVal(29) = RGB(186, 186, 186)      '���_�����m�F���ĂȂ�����Ƃ肠������
    
    ' �t�H���g�F
    ColorValFont(0) = -1
    ColorValFont(1) = RGB(255, 255, 255)
    ColorValFont(2) = RGB(0, 0, 0)
    ColorValFont(3) = RGB(255, 255, 255)
    ColorValFont(4) = RGB(255, 255, 255)
    ColorValFont(5) = RGB(0, 0, 0)
    ColorValFont(6) = RGB(255, 255, 255)
    ColorValFont(7) = RGB(255, 255, 255)
    ColorValFont(8) = RGB(0, 0, 0)
    ColorValFont(9) = RGB(0, 0, 0)
    ColorValFont(10) = RGB(0, 0, 0)
    ColorValFont(11) = RGB(0, 0, 0)
    ColorValFont(12) = RGB(255, 255, 255)
    ColorValFont(13) = RGB(0, 0, 0)
    ColorValFont(14) = RGB(255, 255, 255)
    ColorValFont(15) = RGB(0, 0, 0)
    ColorValFont(16) = RGB(0, 0, 0)
    ColorValFont(17) = RGB(255, 255, 255)
    ColorValFont(18) = RGB(0, 0, 0)
    ColorValFont(19) = RGB(0, 0, 0)
    ColorValFont(20) = RGB(255, 255, 255)
    ColorValFont(21) = RGB(0, 0, 0)
    ColorValFont(22) = RGB(0, 0, 0)
    ColorValFont(23) = RGB(255, 255, 255)
    ColorValFont(24) = RGB(255, 255, 255)
    ColorValFont(25) = RGB(0, 0, 0)
    ColorValFont(26) = RGB(255, 255, 255)
    ColorValFont(27) = RGB(0, 0, 0)
    ColorValFont(28) = RGB(255, 0, 0)
    ColorValFont(29) = RGB(0, 0, 0)
    
    ' �F�R�[�h�i�K�i��̐F�R�[�h�j
    ColorCode(0) = -1   ' ����
    ColorCode(1) = 30   ' ��
    ColorCode(2) = 40   ' ��
    ColorCode(3) = 50   ' ��
    ColorCode(4) = 60   ' ��
    ColorCode(5) = 70   ' ��
    ColorCode(6) = 80   ' ��
    ColorCode(7) = 90   ' ��
    ColorCode(8) = 52   ' �s���N
    ColorCode(9) = 41   ' �D
    ColorCode(10) = 61  ' ��t
    ColorCode(11) = 91  ' ��
    ColorCode(12) = 92  ' ��
    ColorCode(13) = 51  ' �I�����W
    ColorCode(14) = 81  '�`���R���[�g
    ColorCode(15) = 11
    ColorCode(16) = 0
    ColorCode(17) = 21
    ColorCode(18) = 12
    ColorCode(19) = 16
    ColorCode(20) = 17
    ColorCode(21) = 19
    ColorCode(22) = 20
    ColorCode(23) = 22
    ColorCode(24) = 23
    ColorCode(25) = 24
    ColorCode(26) = 25
    ColorCode(27) = 27
    ColorCode(28) = 99
    ColorCode(29) = 99
    
    ' �F�L��
    ColorName(0) = "Notting"    ' ����
    ColorName(1) = "B"   ' ��
    ColorName(2) = "W"   ' ��
    ColorName(3) = "R"   ' ��
    ColorName(4) = "G"   ' ��
    ColorName(5) = "Y"   ' ��
    ColorName(6) = "BR"  ' ��
    ColorName(7) = "L"   ' ��
    ColorName(8) = "P"   ' �s���N
    ColorName(9) = "GY"  ' �D
    ColorName(10) = "LG" ' ��t
    ColorName(11) = "SB" ' ��
    ColorName(12) = "V"  ' ��
    ColorName(13) = "O"  ' �I�����W
    ColorName(14) = "CH" ' �`���R���[�g
    ColorName(15) = "BE"
    ColorName(16) = "ZZ"
    ColorName(17) = "DGY"
    ColorName(18) = "SGY"
    ColorName(19) = "SI"
    ColorName(20) = "DG"
    ColorName(21) = "C"
    ColorName(22) = "TA"
    ColorName(23) = "DL"
    ColorName(24) = "DCH"
    ColorName(25) = "AM"
    ColorName(26) = "VL"
    ColorName(27) = "LA"
    ColorName(28) = "RI" '���炵�F(����Ŏg�p�����)
    ColorName(29) = "LY" '���炵�F(����Ŏg�p�����)
    
End Sub

Function BoxBaseColor(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, i As Long) As String
    
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    ' *** �����`�x�[�X�J���[�`��
    
    ' �����t���O����
    WhiteLineFrg = False
    
    ' �����`�`��
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveCell.left, ActiveCell.Top, ActiveCell.Height, ActiveCell.Height).Select
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 5 * 0.7, 5 * 0.71, 60 * 0.76, 60 * 0.76).Select   '0
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select   '1
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 73 * 0.74, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '2
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 132 * 0.746, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '3
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 194 * 0.745, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '4
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 253 * 0.752, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '5
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 314 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select      '6
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 375 * 0.749, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '7
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 435 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '8
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 496 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '9
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select   '10
    
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (50 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (125 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (185 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (246 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
     
    'xxxx = "12.72.132.193.253.314.375.435.496" '7282-5833
    'xxxx = "50.125.185.246" '7283-2055
    'xxx = Split(xxxx, ".")
    'For Each xx In xxx
        'x = (xx * 0.747) ^ 1.0006
        'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select  'x�̍��W
    'Next xx
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Weight = 1
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    'Selection.OnAction = "���CH"
    ' �x�[�X�F����������
    If filcolor = 1315860 Then
        ' �����t���O�Z�b�g
        WhiteLineFrg = True
        ' ���̐F�𔒂ɕύX
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    End If
    
    '�_�u���ׂ̈̏���_�������O���������m�F

    Selection.Name = �[���} & "_" & cav
    BoxBaseColor = Selection.Name
    
End Function
Function BoxBaseColor2(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, strcolor As Long, i As Long) As String
    
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    ' *** �����`�x�[�X�J���[�`��
    
    ' �����t���O����
    WhiteLineFrg = False
    
    ' �����`�`��
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveCell.left, ActiveCell.Top, ActiveCell.Height, ActiveCell.Height).Select
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 5 * 0.7, 5 * 0.71, 60 * 0.76, 60 * 0.76).Select   '0
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select   '1
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 73 * 0.74, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '2
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 132 * 0.746, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '3
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 194 * 0.745, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '4
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 253 * 0.752, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '5
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 314 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select      '6
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 375 * 0.749, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '7
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 435 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '8
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 496 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '9
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select   '10
    
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (50 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (125 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (185 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (246 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
     
    'xxxx = "12.72.132.193.253.314.375.435.496" '7282-5833
    'xxxx = "50.125.185.246" '7283-2055
    'xxx = Split(xxxx, ".")
    'For Each xx In xxx
        'x = (xx * 0.747) ^ 1.0006
        'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select  'x�̍��W
    'Next xx
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Adjustments.Item(1) = 0.15

    If strcolor = -1 Then
        If filcolor = 13355979 Then '�F�Ă�SI�̎�
            Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(255, 51, 153), 0
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(255, 102, 51), 0.25
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(255, 255, 0), 0.5
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(1, 167, 143), 0.75
            Selection.ShapeRange.Fill.GradientStops.Insert RGB(51, 102, 255), 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
        Else
            'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
            Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.401
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.599
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.99
            Selection.ShapeRange.Fill.GradientStops.Delete 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
        End If
    Else
        Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.401
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.599
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.999
        Selection.ShapeRange.Fill.GradientStops.Delete 1
        Selection.ShapeRange.Fill.GradientStops.Delete 1
    End If
    'Selection.OnAction = "���CH"
    ' �x�[�X�F����������
    If filcolor = 1315860 Then
        ' �����t���O�Z�b�g
        WhiteLineFrg = True
        ' ���̐F�𔒂ɕύX
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(250, 250, 250)
    Else
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    End If
    Selection.ShapeRange.Line.Weight = 1
    
    '�_�u���ׂ̈̏���_�������O���������m�F

    Selection.Name = �[���} & "_" & cav
    BoxBaseColor2 = Selection.Name
    
End Function
Function TerBaseColor2(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, strcolor As Long, i As Long) As String
    
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    ' *** �����`�x�[�X�J���[�`��
    
    ' �����t���O����
    WhiteLineFrg = False
    
    ' �����`�`��
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveCell.left, ActiveCell.Top, ActiveCell.Height, ActiveCell.Height).Select
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 5 * 0.7, 5 * 0.71, 60 * 0.76, 60 * 0.76).Select   '0
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select   '1
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 73 * 0.74, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '2
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 132 * 0.746, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select  '3
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 194 * 0.745, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '4
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 253 * 0.752, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '5
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 314 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select      '6
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 375 * 0.749, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '7
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 435 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '8
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 496 * 0.75, 31 * 0.72, 60 * 0.76, 60 * 0.76).Select     '9
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 12 * 0.7, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select   '10
    
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (50 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (125 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (185 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, (246 * 0.747) ^ 1.0006, 60 * 0.76, 60 * 0.76).Select   '10
     
    'xxxx = "12.72.132.193.253.314.375.435.496" '7282-5833
    'xxxx = "50.125.185.246" '7283-2055
    'xxx = Split(xxxx, ".")
    'For Each xx In xxx
        'x = (xx * 0.747) ^ 1.0006
        'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (4 * 0.747) ^ 1.0006, 99 * 0.745, 60 * 0.76, 60 * 0.76).Select  'x�̍��W
    'Next xx
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Adjustments.Item(1) = 0.15
    If strcolor = -1 Then
        'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
        Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.401
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.599
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.99
        Selection.ShapeRange.Fill.GradientStops.Delete 1
        Selection.ShapeRange.Fill.GradientStops.Delete 1
    Else
        Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.401
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.599
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.999
        Selection.ShapeRange.Fill.GradientStops.Delete 1
        Selection.ShapeRange.Fill.GradientStops.Delete 1
    End If
    
    Selection.ShapeRange.Line.Weight = 1
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    'Selection.OnAction = "���CH"
    ' �x�[�X�F����������
    If filcolor = 1315860 Then
        ' �����t���O�Z�b�g
        WhiteLineFrg = True
        ' ���̐F�𔒂ɕύX
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    End If
    
    '�_�u���ׂ̈̏���_�������O���������m�F

    Selection.Name = �[���} & "_" & cav
    TerBaseColor2 = Selection.Name
    
End Function


Function BoxStrColor(i As Long, ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long) As String
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    ' *** �����`�X�g���C�v�J���[�`��
    xLeft = (xLeft * 0.747) ^ 1.0006
    yTop = (yTop * 0.747) ^ 1.0006
    myWidth = (myWidth * 0.747) ^ 1.0006
    myHeight = (myHeight * 0.747) ^ 1.0006

    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, xLeft + myWidth * 0.011 + myWidth * 0.8, yTop + myHeight * 0.01)
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.993, yTop + myHeight * 0.01
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.993, yTop + myHeight * 0.01 + myHeight * 0.28
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.28, yTop + myHeight * -0.01 + myHeight
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.011, yTop + myHeight * -0.011 + myHeight
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.011, yTop + myHeight * -0.011 + myHeight * 0.8
        .AddNodes msoSegmentLine, msoEditingAuto, xLeft + myWidth * 0.011 + myWidth * 0.8, yTop + myHeight * 0.005
        .ConvertToShape.Select
    End With
    
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Visible = msoFalse
    '���̐F�𔒂ɕύX
    If WhiteLineFrg = True Then Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    
    '�I�[�g�V�F�C�v�̖��O��Ԃ�
    Selection.Name = �[���} & "_" & cav & "_str"
    BoxStrColor = Selection.Name
        
End Function

Sub DeleteAllShapes()
    
    ' *** ���[�N�V�[�g��̂��ׂĂ̐}�`����������
    
    ' �G���[���������Ă��X�L�b�v����
    On Error Resume Next
    
    ' �A�N�e�B�u�ȃ��[�N�V�[�g�̃I�[�g�V�F�C�v���̐}�`������
    ActiveSheet.Shapes.SelectAll
    Selection.ShapeRange.Delete
    
    ' �G���[�̃X�L�b�v������
    On Error GoTo 0
    
End Sub

Public Function BoxFill(xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, _
                 FilColor1 As Variant, i As Long, Optional FilColor2 As Variant = 0, Optional �}���}1, _
                 Optional �V�[���h�t���O As String, Optional �I���o�� As String, Optional ByVal �T�C�Y�� As String, _
                 Optional �n�� As String) As String
    
    ' *** �l�p�`�`��
    ' �ϐ��̐錾
    Dim BaseName As String      ' �x�[�X�F�̃I�u�W�F�N�g��
    Dim StrName As String       ' �X�g���C�v�F�̃I�u�W�F�N�g��
    Dim clocode1 As Long        ' �F�P�i�[�p
    Dim clocode2 As Long        ' �F�Q�i�[�p
    Dim CloCode3 As Long
    Dim CloCodeFont1 As Long    ' CloCode1�ɑ΂���t�H���g�F
    Dim BufSize As Single       ' �T�C�Y���ۑ��p
    Dim sFontSize As Long
    Dim myFontColor As Long
    Dim baseSize As Single

    �n��s = Split(�n��, "!")
    If InStr(�n��s(5), "��") > 0 Then
        �� = "Au" & vbLf
        ��a = 4
    Else
        �� = ""
        ��a = 1
    End If
    �F�� = FilColor1
    If FilColor2 <> 0 Then �F�� = �F�� & "/" & FilColor2
    Call �F�ϊ�(�F��, clocode1, clocode2, clofont)
    
    'If �F�Ŕ��f = True Then clocode1 = 16777215: clofont = 0
    If �}���}1 <> "" Then Call �F�ϊ�(�}���}1, CloCode3, 0, 0)
    
    ' �T�C�Y�l�擾
    BufSize = Size
    
    ' �T�C�Y���P�D�Q�ȉ��̏ꍇ�͕␳����(Excel2000�p�̃o�O�΍�)
    If BufSize < 1.2 Then BufSize = 2
    BaseName = BoxBaseColor2(xLeft, yTop, myWidth, myHeight, clocode1, clocode2, i)
    '�e�L�X�g��}�`����͂ݏo���ĕ\������
    Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
    Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
    Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
    Selection.Font.Name = myFont
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Selection.ShapeRange.TextFrame2.MarginLeft = 0
    Selection.ShapeRange.TextFrame2.MarginRight = 0
    Selection.ShapeRange.TextFrame2.MarginTop = 0
    Selection.ShapeRange.TextFrame2.MarginBottom = 0
    Select Case �n���}�^�C�v
    Case "�`�F�b�J�[�p", "��H����", "�\��", "����[��"
        If InStr(�I���o��, "!") = 0 Then
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = �� & �I���o�� & vbLf & �T�C�Y��
            '���[��n���Ȃ�1�s�ڂɃA���_�[�o�[
            If �n��s(1) = "1" Then
                Selection.ShapeRange.TextFrame2.TextRange.Characters(��a, Len(�I���o��)).Font.UnderlineStyle = msoUnderlineSingleLine
            End If
            '���[�������[�q�Ȃ�1�s�ڂ��Α�
            If �n��s(2) = "1" Then
                Selection.ShapeRange.TextFrame2.TextRange.Characters(��a, Len(�I���o��)).Font.Italic = msoTrue
            End If
        Else
            �I���o��A = Split(�I���o��, "!")
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = �� & �I���o��A(0) & vbLf & �I���o��A(1)
            �I���o�� = �I���o��A(0)
        End If
    
        If Selection.Width > Selection.Height Then
            sFontSize = Selection.Height * 0.48
            gyospace = 0.8
        Else
            sFontSize = Selection.Width * 0.48
            gyospace = 0.8
        End If
        If Len(�I���o��) = 4 Then
            sFontSize = sFontSize * 0.87
        End If
        
        myFontColor = clofont
        '�X�g���C�v�͌��ʂ��g��
        If clocode1 <> clocode2 Or �F�Ŕ��f = True Then
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                If �F�Ŕ��f = True Then
                    .color = 16777215 '��
                    .Radius = 11
                Else
                    .color = clocode1
                    .Radius = 8
                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = clofont
                End If
                .color.TintAndShade = 0
                .color.Brightness = 0
                .Transparency = 0#
            End With
        Else
            Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = clofont
        End If
        
        'Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = &HFFFFFF And Not Selection.ShapeRange.Fill.ForeColor.RGB '�t�H���g�𔽑ΐF�ɂ���
    
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 4).ParagraphFormat.SpaceWithin = 0.1 '�s��
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = gyospace  '�s��
    Case Else
        If Len(�T�C�Y��) > 3 Then �T�C�Y��t = Left(�T�C�Y��, 3) Else �T�C�Y��t = �T�C�Y��
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = �� & �T�C�Y��t
       
        myFontColor = clofont
        '�X�g���C�v�͌��ʂ��g��
        If clocode1 <> clocode2 Or �F�Ŕ��f = True Then
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                If �F�Ŕ��f = True Then
                    .color = 16777215 '��
                    .Radius = 10
                Else
                    .color = clocode1
                    .Radius = 8
                End If
                .color.TintAndShade = 0
                .color.Brightness = 0
                .Transparency = 0#
            End With
        End If
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        'Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = &HFFFFFF And Not Selection.ShapeRange.Fill.ForeColor.RGB '�t�H���g�𔽑ΐF�ɂ���

        If Len(�T�C�Y��t) < Len(��) - 2 Then �T�C�Y��t = "Au"
        Select Case Len(�T�C�Y��t)
            Case 1
                baseSize = 0.8
            Case 2
                baseSize = 0.9
            Case Else
                baseSize = 0.9
        End Select
        
        If Selection.Width > Selection.Height Then
            sFontSize = Selection.Height * (1.6 / Len(Replace(�T�C�Y��t, ".", ""))) * baseSize
            If InStr(�T�C�Y��t, ".") > 0 Then sFontSize = sFontSize - (Selection.Height * 0.3)
        Else
            sFontSize = Selection.Width * (1.6 / Len(Replace(�T�C�Y��t, ".", ""))) * baseSize
            If InStr(�T�C�Y��t, ".") > 0 Then sFontSize = sFontSize - (Selection.Width * 0.3)
        End If
        
    End Select
    '��ɒ[���o�H�ŕ��������������班���傫������
    If Len(�I���o��) = 4 Then
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
    Else
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize * 1.2
    End If
        
    
    If �� <> "" Then
        myLen = Len(Selection.ShapeRange.TextFrame2.TextRange.Characters.Text)
        Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 2).Font.Size = sFontSize * 1
        gyospace = 0.7
        Selection.ShapeRange.TextFrame2.TextRange.Characters(1, ��a).ParagraphFormat.SpaceWithin = gyospace
    End If
    
'    ���C�� = (Selection.Width + Selection.Height) / 60
'    If ���C�� < 0.25 Then ���C�� = 0.25
'    If ���C�� > 2 Then ���C�� = 2
    'line.weight��1�ɌŒ�(���ވꗗ+�̃n���}�쐬�Ɠ����ɂ���)
    Selection.ShapeRange.Line.Weight = 1
    
    �n��s = Split(�n��, "!")
    
    If �F�Ŕ��f = True Then
        For i2 = 1 To UBound(�n���F�ݒ�, 2)
            If �n��s(0) = �n���F�ݒ�(2, i2) Then
                Selection.Font.color = �n���F�ݒ�(1, i2)
                Exit For
            End If
        Next i2
    End If
    
    If �n��s(4) <> "" Then
        If ��d�W�~flg = True Then
            If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
            Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 4
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
        End If
    End If
        
    If �n����ƕ\�� <> "" Then
        If CLng(�n��s(3)) < CLng(�n����ƕ\��) Then
            Select Case �n���\��
            Case "1"
                Call ���CH_��n��
            Case "2"
                Call ���CH_��n��_2
            Case "3"
                Call ���CH_��n��_3
            End Select
        End If
    Else
        If �n��s(0) = "��n��" Then
            Select Case �n���\��
            Case "1"
                Call ���CH_��n��
            Case "2"
                Call ���CH_��n��_2
            Case "3"
                Call ���CH_��n��_3
            End Select
        End If
    End If
    
    If �}���}1 <> "" Then
        Set obtemp = Selection.ShapeRange
        StrName = BoxFeltTip(xLeft, yTop, myWidth, myHeight, CloCode3, clocode1, myFontColor)
        '�O���[�v��
        obtemp.Select False
        'ActiveSheet.Shapes.Range(BaseName).Select False
        Selection.Group.Select
        Selection.Name = �[���} & "_" & cav & "_g"
    End If
    
    '�����Ɠ������O��2�ȏ㖳�����m�F_��������_�u��
    Dim �������O�̐� As Long: �������O�̐� = 0
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        If objShp.Name = �[���} & "_" & cav Or objShp.Name = �[���} & "_" & cav & "_g" Then
            objShp.Select
            �������O�̐� = �������O�̐� + 1
        End If
    Next
    
    '�_�u���̉摜�T�C�Y�ύX
    If �������O�̐� > 1 Then
        Dim �_�u��1�{�� As Long: �_�u��1�{�� = 0
        For Each objShp In ActiveSheet.Shapes
            If objShp.Name = Selection.Name Then
                '2�s�ڂ��|�C���g���ǂ������f
                On Error Resume Next
                zz = objShp.TextFrame2.TextRange.Characters.Text
                If Err <> 0 Then '�}���}�ƃO���[�v�����Ă���ꍇ
                    'objShp.Ungroup
                    For Each objShp2 In objShp.GroupItems
                        'ActiveSheet.Shapes.Range(�[���} & "_" & cav & "_g").Ungroup
                        'ActiveSheet.Shapes.Range(�[���} & "_" & CAV).Ungroup
                        If Not objShp2.Name Like "*_Felt" Then
                            zz = objShp2.TextFrame2.TextRange.Characters.Text
                            Set objShp3 = objShp2
                        End If
                    Next
                Else
                    Set objShp3 = objShp
                End If
                On Error GoTo 0
Return0:
                zzz = InStr(zz, vbLf)
                cc = ""
                If zzz > 0 Then
                    aa = Left(zz, zzz - 1)
                    bb = Replace(Mid(zz, zzz), vbLf, "")
                    If Len(bb) > 3 Then cc = 1 Else cc = 0
                End If
                If �_�u��1�{�� = 0 Then
                    objShp.Height = objShp.Height / 2
                    �_�u��1�{�� = 1
                    objShp.Name = �[���} & "_" & cav & "_w1"
                    If cc = "" Then
                        objShp3.TextFrame2.TextRange.Characters.Font.Size = objShp3.TextFrame2.TextRange.Characters.Font.Size / 2
                        objShp3.TextFrame2.TextRange.Characters.Text = zz
                    Else
                        objShp3.TextFrame2.TextRange.Characters.Text = aa
                    End If
                Else
                    objShp.Height = objShp.Height / 2
                    objShp.Top = objShp.Top + objShp.Height
                    objShp.Name = �[���} & "_" & cav & "_w2"
                    If cc = "" Then
                        objShp3.TextFrame2.TextRange.Characters.Font.Size = objShp3.TextFrame2.TextRange.Characters.Font.Size / 2
                        objShp3.TextFrame2.TextRange.Characters.Text = zz
                    Else
                        objShp3.TextFrame2.TextRange.Characters.Text = aa
                    End If
                End If
            End If
        Next
        ActiveSheet.Shapes.Range(�[���} & "_" & cav & "_w1").Select False
        Selection.Group.Select
        Selection.Name = �[���} & "_" & cav
    End If

    If �[��cav�W�� = "" Then
        �[��cav�W�� = Selection.Name
    Else
        �[��cav�W�� = �[��cav�W�� & "," & Selection.Name
    End If
Exit Function

End Function

Function TerFill(xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, _
                 FilColor1 As Variant, i As Long, Optional FilColor2 As Variant = 0, Optional �}���}1, Optional �V�[���h�t���O As String) As String
    
    ' *** �l�p�`�`��
    ' �ϐ��̐錾
    Dim BaseName As String      ' �x�[�X�F�̃I�u�W�F�N�g��
    Dim StrName As String       ' �X�g���C�v�F�̃I�u�W�F�N�g��
    Dim clocode1 As Long        ' �F�P�i�[�p
    Dim clocode2 As Long        ' �F�Q�i�[�p
    Dim CloCode3 As Long
    Dim BufSize As Single       ' �T�C�Y���ۑ��p
    Dim sFontSize As Long
    
    �F�� = FilColor1
    If FilColor2 <> 0 Then �F�� = �F�� & "/" & FilColor2
    Call �F�ϊ�(�F��, clocode1, clocode2, clofont)
    
    If �}���}1 <> "" Then
        Call �F�ϊ�(�}���}1, CloCode3, 0, 0)
    Else
        CloCode3 = -1
    End If
    
    ' �T�C�Y�l�擾
    BufSize = Size
    
    ' �T�C�Y���P�D�Q�ȉ��̏ꍇ�͕␳����(Excel2000�p�̃o�O�΍�)
    If BufSize < 1.2 Then BufSize = 2
    ' �i������ƈꌾ�j
    ' Excel2000���Ă΃q�h���񂾂��āI
    ' �}�N���̃R�[�h����ɏ��T�C�Y�̃t���[�t�H�[����
    ' �쐬���悤�Ƃ���ƃG���[�N�����񂾂��Ă΁I
    ' �u�}�N���̋L�^�v�����Ȃ���ɏ��T�C�Y��
    ' �t���[�t�H�[����`�悷��͉̂�����Ȃ��̂�
    ' �L�^���ꂽ�R�[�h�����s���ăt���[�t�H�[����
    ' �`�悵�悤�Ƃ���Ɓu�I�[�g���[�V�����G���[�v
    ' �Ȃ�Ă̂���������́I�I���悱��I�H
    ' ��
    ' ...
    'CloCode1
    'BaseName = BoxBaseColor(xLeft, yTop, myWidth, myHeight, CloCode1, i)
    BaseName = TerBaseColor2(xLeft, yTop, myWidth, myHeight, clocode1, clocode2, i)
    
    myFontColor = CloCodeFont1 '�t�H���g�F���x�[�X�F�Ō��߂�
    '�X�g���C�v�͌��ʂ��g��
    If clocode2 <> clocode1 Then
        With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
            .color = clocode1
            .color.TintAndShade = 0
            .color.Brightness = 0
            .Transparency = 0#
            .Radius = 8
        End With
    End If
'If �V�[���h�t���O = "S" Then
'    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = "S"
'    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
'    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
'    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
'    Selection.ShapeRange.TextFrame2.MarginLeft = 0
'    Selection.ShapeRange.TextFrame2.MarginRight = 0
'    Selection.ShapeRange.TextFrame2.MarginTop = 0
'    Selection.ShapeRange.TextFrame2.MarginBottom = 0
'    If Selection.Width > Selection.Height Then
'        sFontSize = Selection.Height * 1.5
'    Else
'        sFontSize = Selection.Width * 1.5
'    End If
'    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
'    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = &HFFFFFF And Not Selection.ShapeRange.Fill.ForeColor.RGB
'End If
    'CloCode2
'    If CloCode2 >= 0 And CloCode1 <> CloCode2 Then
'        StrName = BoxStrColor(i, xLeft, yTop, myWidth, myHeight, CloCode2)
'        '�O���[�v��
'        ActiveSheet.Shapes.Range(BaseName).Select False
'        Selection.Group.Select
'        'ActiveSheet.Shapes.Range(Array(BaseName, StrName)).Group.Select
'        Selection.Name = �[���} & "_" & Cav
'        BaseName = Selection.Name
'    End If
    'CloCode
    If CloCode3 >= 0 Then
        StrName = BoxFeltTip(xLeft, yTop, myWidth, myHeight, CloCode3, clocode1, myFontColor)
        '�O���[�v��
        ActiveSheet.Shapes.Range(BaseName).Select False
        Selection.Group.Select
        Selection.Name = �[���} & "_" & cav
    End If
    
    '�����Ɠ������O��2�ȏ㖳�����m�F_��������_�u��
    Dim �������O�̐� As Long: �������O�̐� = 0
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        If objShp.Name = Selection.Name Then
            �������O�̐� = �������O�̐� + 1
        End If
    Next
    
    '�_�u���̉摜�T�C�Y�ύX
    If �������O�̐� > 1 Then
        Dim �_�u��1�{�� As Long: �_�u��1�{�� = 0
        For Each objShp In ActiveSheet.Shapes
            If objShp.Name = Selection.Name Then
                If �_�u��1�{�� = 0 Then
                    objShp.Width = objShp.Width / 2
                    objShp.Name = Selection.Name & "_w1"
                    �_�u��1�{�� = �_�u��1�{�� + 1
                ElseIf �_�u��1�{�� = 1 Then
                    objShp.Width = objShp.Width / 2
                    objShp.Left = objShp.Left + objShp.Width
                    objShp.Name = Selection.Name & "_w2"
                    �_�u��1�{�� = �_�u��1�{�� + 1
                Else
                    Stop
                End If
            End If
        Next
        ActiveSheet.Shapes.Range(�[���} & "_" & cav & "_w1").Select False
        Selection.Group.Select
        Selection.Name = �[���} & "_" & cav
    End If
        
    'line.weight��1�ɌŒ�(���ވꗗ+�̃n���}�쐬�Ɠ����ɂ���)
    Selection.ShapeRange.Line.Weight = 1
    
    '�[�q�Ȃ̂œd�����Ŕw�ʂɈړ�
    Selection.ShapeRange.ZOrder msoSendToBack
    
    
    If �[��cav�W�� = "" Then
        �[��cav�W�� = �[���} & "_" & cav
    Else
        �[��cav�W�� = �[��cav�W�� & "," & �[���} & "_" & cav
    End If
    ' �쐬���ꂽ�I�[�g�V�F�C�v�̖��O��Ԃ�
    'ActiveSheet.Shapes.Range(�[���}).Select False
    'Selection.Group.Select
    'Selection.Name = �[��
    'Set target = Union(ActiveSheet.Shapes.Range(�[��), ActiveSheet.Shapes.Range(BoxFill))
    'ActiveSheet.Shapes.Range(Array(�[��, BoxFill)).Group.Select
    'Selection.Name = �[���}
    'BoxFill = Selection.Name
    
End Function


Function BonFill(xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, Optional RowStr) As String
    
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    
    ' *** �l�p�`�`��
    ' �ϐ��̐錾
    Dim BaseName As String      ' �x�[�X�F�̃I�u�W�F�N�g��
    Dim StrName As String       ' �X�g���C�v�F�̃I�u�W�F�N�g��
    Dim clocode1 As Long        ' �F�P�i�[�p
    Dim clocode2 As Long        ' �F�Q�i�[�p
    Dim CloCode3 As Long
    Dim BufSize As Single       ' �T�C�Y���ۑ��p
    Dim sFontSize As Long


    Dim filc As String, strc As String
    ' �f�[�^�̏�����������Ă��Ȃ����͏���������
    
    '�T�C�Y�l�擾
    BufSize = Size
    
    '�T�C�Y��1.2�ȉ��̏ꍇ�͕␳����(Excel2000�p�̃o�O�΍�)
    If BufSize < 1.2 Then BufSize = 2
    
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Adjustments.Item(1) = 0.05
    Selection.ShapeRange.Fill.OneColorGradient 1, 1, 1
    
    a = 1 / (UBound(RowStr) + 1)
    For q = LBound(RowStr) To UBound(RowStr)
        V = Split(RowStr(q), "_")
        Call �F�ϊ�(V(4), filcolor, strcolor, fontcolor)
        If q = LBound(ColorCode) Then Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0 / (UBound(RowStr) + 1)
'        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.05 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.3 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.301 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.699 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.7 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.99 * a + (q / (UBound(RowStr) + 1))
        Selection.ShapeRange.Fill.GradientStops.Insert fontcolor, 0.999 * a + (q / (UBound(RowStr) + 1))
    Next q
    Selection.ShapeRange.Fill.GradientStops.Delete 1
    Selection.ShapeRange.Fill.GradientStops.Delete 1

    Selection.ShapeRange.Line.Weight = 1
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
        
    '�t�H���g�F���x�[�X�F�Ō��߂�
    myFontColor = fontcolor
    '�X�g���C�v�͌��ʂ��g��
    If clocode1 = clocode2 Then
        With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
            .color = clocode1
            .color.TintAndShade = 0
            .color.Brightness = 0
            .Transparency = 0#
            .Radius = 8
        End With
    End If
        
    'line�̃T�C�Y�ύX
    Selection.ShapeRange.Line.Weight = 1
    
    'bonda�Ȃ̂œd�����Ŕw�ʂɈړ�
    Selection.ShapeRange.ZOrder msoSendToBack
    
    Selection.Name = �[���} & "_" & cav
    
    If �[��cav�W�� = "" Then
        �[��cav�W�� = �[���} & "_" & cav
    Else
        �[��cav�W�� = �[��cav�W�� & "," & �[���} & "_" & cav
    End If
    
End Function

Function CircleBaseColor(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, i As Long) As String
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    ' �����t���O����
    WhiteLineFrg = False
    
    ' ���~�`�`�恕�F�ݒ�
    ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Weight = 1
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    'Selection.OnAction = "���CH"
    ' �x�[�X�F����������
    If filcolor = 1315860 Then
        ' �����t���O�Z�b�g
        WhiteLineFrg = True
        ' ���̐F�𔒂ɕύX
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    End If
    
    ' �I�[�g�V�F�C�v�̖��O��Ԃ�
    Selection.Name = �[���} & "_" & cav
    CircleBaseColor = Selection.Name
   
End Function

Function CircleBaseColor2(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, strcolor As Long, i As Long) As String
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    ' �����t���O����
    WhiteLineFrg = False
    
    ' ���~�`�`�恕�F�ݒ�
    ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
'    If �F�Ŕ��f = True Then
'
'    Else
        If strcolor = -1 Then
            'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
            Selection.ShapeRange.Fill.OneColorGradient 3, 1, 1
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.401
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.599
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.99
            Selection.ShapeRange.Fill.GradientStops.Delete 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
        Else
            Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.4
            Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.401
            Selection.ShapeRange.Fill.GradientStops.Insert strcolor, 0.599
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.6
            Selection.ShapeRange.Fill.GradientStops.Insert filcolor, 0.999
            Selection.ShapeRange.Fill.GradientStops.Delete 1
            Selection.ShapeRange.Fill.GradientStops.Delete 1
        End If
        ' �x�[�X�F����������
        If filcolor = 1315860 Then
            ' �����t���O�Z�b�g
            WhiteLineFrg = True
            ' ���̐F�𔒂ɕύX
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
        Else
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
        End If
'    End If
    Selection.ShapeRange.Line.Weight = 1
    'Selection.OnAction = "���CH"

    ' �I�[�g�V�F�C�v�̖��O��Ԃ�
    Selection.Name = �[���} & "_" & cav
    CircleBaseColor2 = Selection.Name
   
End Function


Function CircleStrColor(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, i As Long) As String
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    ' *** ���~�`�X�g���C�v�J���[
    Dim �Z�� As Object
    Set �Z�� = ActiveCell
    
    xLeft = (xLeft * 0.747) ^ 1.0006
    yTop = (yTop * 0.747) ^ 1.0006
    myWidth = (myWidth * 0.747) ^ 1.0006
    myHeight = (myHeight * 0.747) ^ 1.0006
    
    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, xLeft + myWidth * 0.733, yTop + myHeight * 0.068)
        .AddNodes msoSegmentCurve, msoEditingCorner, xLeft + myWidth * 0.818, yTop + myHeight * 0.115, _
                                                     xLeft + myWidth * 0.887, yTop + myHeight * 0.185, _
                                                     xLeft + myWidth * 0.933, yTop + myHeight * 0.265
        .AddNodes msoSegmentCurve, msoEditingCorner, xLeft + myWidth * 0.265, yTop + myHeight * 0.933
        .AddNodes msoSegmentCurve, msoEditingCorner, xLeft + myWidth * 0.185, yTop + myHeight * 0.887, _
                                                     xLeft + myWidth * 0.115, yTop + myHeight * 0.818, _
                                                     xLeft + myWidth * 0.068, yTop + myHeight * 0.733
        .AddNodes msoSegmentCurve, msoEditingCorner, xLeft + myWidth * 0.733, yTop + myHeight * 0.068
        .ConvertToShape.Select
    End With
    
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Visible = msoFalse
    
    '���̐F�𔒂ɕύX
    If WhiteLineFrg = True Then Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
    
    ' �I�[�g�V�F�C�v�̖��O��Ԃ�
    Selection.Name = �[���} & "_" & cav & "_str"
    CircleStrColor = Selection.Name
    
End Function
Function BoxNull(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, _
                 ByVal �I���o�� As String, ByVal �T�C�Y�� As String, ByVal EmptyPlug As String, _
                    ByVal PlugColor As String) As String
    ' �ϐ��̐錾
    Dim BaseName As String      ' �x�[�X�F�̃I�u�W�F�N�g��
    Dim BufSize As Single       ' �T�C�Y���ۑ��p
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
    
    '��������鎞
    If EmptyPlug <> "" Then
        'ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = "*"
        'Selection.Font.Name = myFont
        sFontSize = Selection.Width * 1.2
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
        Call �F�ϊ�(PlugColor, clocode1, clocode2, clofont)
        myFontColor = clocode1
         Selection.ShapeRange.TextFrame2.MarginLeft = 0
         Selection.ShapeRange.TextFrame2.MarginRight = 0
         Selection.ShapeRange.TextFrame2.MarginTop = 0
         Selection.ShapeRange.TextFrame2.MarginBottom = 0
         Selection.ShapeRange.TextFrame2.Orientation = msoTextOrientationHorizontalRotatedFarEast
         Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
         Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
         Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
         Selection.ShapeRange.TextFrame2.TextRange.Font.Line.Visible = True
         Selection.ShapeRange.TextFrame2.TextRange.Font.Line.ForeColor.RGB = 0
    Else '��
        If InStr(�I���o��, "!") = 0 Then
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = �I���o�� & vbCrLf & " "
            If Selection.Width > Selection.Height Then
                sFontSize = Selection.Height * 0.4
                gyospace = 0.8
            Else
                sFontSize = Selection.Width * 0.4
                gyospace = 0.8
            End If
        Else
            �I���o��A = Split(�I���o��, "!")
            Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = �I���o��A(0) & vbCrLf & �I���o��A(1)
            �I���o�� = �I���o��A(0)
            If Selection.Width > Selection.Height Then
                sFontSize = Selection.Height * 0.4
                gyospace = 0.8
            Else
                sFontSize = Selection.Width * 0.4
                gyospace = 0.8
            End If
        End If
        If Len(�I���o��) = 4 Then sFontSize = sFontSize * 0.87
        Selection.Font.Name = myFont
        Selection.ShapeRange.TextFrame2.TextRange.Characters(Len(�I���o��) + 1, Len(" ") + 1).ParagraphFormat.SpaceWithin = gyospace  '�s��
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
        Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
        Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        Selection.ShapeRange.TextFrame2.MarginLeft = 0
        Selection.ShapeRange.TextFrame2.MarginRight = 0
        Selection.ShapeRange.TextFrame2.MarginTop = 0
        Selection.ShapeRange.TextFrame2.MarginBottom = 0
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        Selection.ShapeRange.Fill.Patterned msoPatternWideDownwardDiagonal
        Selection.ShapeRange.Fill.BackColor.RGB = RGB(230, 230, 230)
    End If
           
    'line�̃T�C�Y�ύX
    ���C�� = (Selection.Width + Selection.Height) / 60
    If ���C�� < 0.25 Then ���C�� = 0.25
    If ���C�� > 2 Then ���C�� = 2
    Selection.Name = �[���} & "_" & cav & ""
    Name1 = Selection.Name
        
    Selection.ShapeRange.Line.Weight = ���C��
    
    
    If �t�H�[������̌Ăяo�� = False Then
        If ��d�W�~flg = True Then
            If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
            Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 4
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
        End If
    End If
        
    If �[��cav�W�� = "" Then
        �[��cav�W�� = �[���} & "_" & cav
    Else
        �[��cav�W�� = �[��cav�W�� & "," & �[���} & "_" & cav
    End If
    
End Function
Function CircleNull(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, _
                    ByVal �I���o�� As String, ByVal �T�C�Y�� As String, ByVal EmptyPlug As String, _
                    ByVal PlugColor As String) As String

    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    ' �ϐ��̐錾
    Dim BaseName As String      ' �x�[�X�F�̃I�u�W�F�N�g��
    Dim BufSize As Single       ' �T�C�Y���ۑ��p
    
    ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
        
    If �n���}�^�C�v = "�`�F�b�J�[�p" And �I���o�� <> "" Then
        '�t�H���g�T�C�Y�����߂�
        If Selection.Width > Selection.Height Then
            sFontSize = Selection.Height * 0.32
            gyospace = 0.7
        Else
            sFontSize = Selection.Width * 0.32
            gyospace = 0.8
        End If
        If Len(�I���o��) = 4 Then
            sFontSize = sFontSize * 0.87
        End If
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = �I���o�� & vbLf & " "
        'Selection.Font.Name = myFont
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
        Selection.ShapeRange.TextFrame2.TextRange.Characters(Len(�I���o��) + 1, Len(" ") + 1).ParagraphFormat.SpaceWithin = gyospace  '�s��
    End If
    
    If �n���}�^�C�v <> "�`�F�b�J�[�p" And CStr(�n���\��) <> "4" Then
       'ActiveSheet.Shapes.AddShape(msoShapeOval, (xLeft * 0.747) ^ 1.0006, (yTop * 0.747) ^ 1.0006, (myWidth * 0.747) ^ 1.0006, (myHeight * 0.747) ^ 1.0006).Select
       Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = "*"
       'Selection.Font.Name = myFont
       sFontSize = Selection.Width * 1.2
       Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
       'Selection.ShapeRange.Fill.UserPicture "D:\18_���ވꗗ\���ވꗗ�쐬�V�X�e��_�p�[�c\NullCircle.png"
       Call �F�ϊ�(PlugColor, clocode1, clocode2, clofont)
       myFontColor = clocode1
        Selection.ShapeRange.TextFrame2.MarginLeft = 0
        Selection.ShapeRange.TextFrame2.MarginRight = 0
        Selection.ShapeRange.TextFrame2.MarginTop = 0
        Selection.ShapeRange.TextFrame2.MarginBottom = 0
        Selection.ShapeRange.TextFrame2.Orientation = msoTextOrientationHorizontalRotatedFarEast
        Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
        Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        
        Selection.ShapeRange.TextFrame2.TextRange.Font.Line.Visible = True
        Selection.ShapeRange.TextFrame2.TextRange.Font.Line.ForeColor.RGB = clofont
        myred = myFontColor Mod 256
        myGreen = Int(myFontColor / 256) Mod 256
        myBlue = Int(myFontColor / 256 / 256)
        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(myred * 1.1, myGreen * 1.1, myBlue * 1.1)
    Else
        Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
        Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        Selection.ShapeRange.TextFrame2.MarginLeft = 0
        Selection.ShapeRange.TextFrame2.MarginRight = 0
        Selection.ShapeRange.TextFrame2.MarginTop = 0
        Selection.ShapeRange.TextFrame2.MarginBottom = 0
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        Selection.ShapeRange.Fill.Patterned msoPatternWideDownwardDiagonal
        Selection.ShapeRange.Fill.BackColor.RGB = RGB(220, 220, 220)
    End If
    'line�̃T�C�Y�ύX
'    ���C�� = (Selection.Width + Selection.Height) / 60
'    If ���C�� < 0.25 Then ���C�� = 0.25
'    If ���C�� > 2 Then ���C�� = 2
    Selection.ShapeRange.Line.Weight = 1
'    Selection.ShapeRange.Line.Weight = ���C��
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(20, 20, 20)
    Selection.Name = �[���} & "_" & cav
    
    xLeft = (xLeft * 0.747) ^ 1.0006
    yTop = (yTop * 0.747) ^ 1.0006
    myWidth = (myWidth * 0.747) ^ 1.0006
    myHeight = (myHeight * 0.747) ^ 1.0006
    
    If �[��cav�W�� = "" Then
        �[��cav�W�� = �[���} & "_" & cav
    Else
        �[��cav�W�� = �[��cav�W�� & "," & �[���} & "_" & cav
    End If
    
    If �t�H�[������̌Ăяo�� = False Then
        If ��d�W�~flg = True Then
            If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
            Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 4
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
        End If
    End If

    '���i�Ԃ̕\�L
    If EmptyPlug <> "" Then
        ���ǉ�flg = 0
        For a = 0 To ���c
            If ���\�L(0, a) = EmptyPlug Then
                ���\�L(1, a) = ���\�L(1, a) + 1
                ���ǉ�flg = 1
            End If
        Next a
        If ���ǉ�flg = 0 Then
            ���c = ���c + 1
            ReDim Preserve ���\�L(2, ���c)
            ���\�L(0, ���c) = EmptyPlug
            ���\�L(1, ���c) = 1
            ���\�L(2, ���c) = PlugColor
        End If
    End If
End Function
Public Function CircleFill(i As Long, xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, _
                 FilColor1 As Variant, Optional FilColor2 As Variant = 0, Optional �}���}1 As String, _
                 Optional �V�[���h�t���O As String, Optional �I���o�� As String, Optional ByVal �T�C�Y�� As String, _
                 Optional �n�� As String) As String
    ' �ϐ��̐錾
    Dim BaseName As String      ' �x�[�X�F�̃I�u�W�F�N�g��
    Dim StrName As String       ' �X�g���C�v�F�̃I�u�W�F�N�g��
    Dim clocode1 As Long        ' �F�P�i�[�p
    Dim clocode2 As Long        ' �F�Q�i�[�p
    Dim CloCode3 As Long
    Dim CloCodeFont1 As Long    ' CloCode1�ɑ΂���t�H���g�F
    Dim BufSize As Single       ' �T�C�Y���ۑ��p
    
    �n��s = Split(�n��, "!")
    
    If InStr(�n��s(5), "��") > 0 Then
        �� = "Au" & vbLf
        ��a = 4
    Else
        �� = ""
        ��a = 1
    End If
    �F�� = FilColor1
    If FilColor2 <> 0 Then �F�� = �F�� & "/" & FilColor2
    Call �F�ϊ�(�F��, clocode1, clocode2, clofont)
    
    'If �F�Ŕ��f = True Then clocode1 = 16777215: clofont = 0
    
    If �}���}1 <> "" Then Call �F�ϊ�(�}���}1, CloCode3, 0, 0)
    
    BufSize = Size
    
    ' �T�C�Y��100�ȉ��̏ꍇ�͕␳����(���T�C�Y�ł̓t���[�t�H�[���ł̌덷���傫������)
    'If BufSize < 100 Then BufSize = 100
    
    ' �x�[�X�F�`��
    'BaseName = CircleBaseColor(xLeft, yTop, myWidth, myHeight, CloCode1, i)
    BaseName = CircleBaseColor2(xLeft, yTop, myWidth, myHeight, clocode1, clocode2, i)
    '�e�L�X�g��}�`����͂ݏo���ĕ\������
    Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
    Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
    Selection.ShapeRange.TextFrame2.WordWrap = msgfalse
    Selection.Font.Name = myFont
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Selection.ShapeRange.TextFrame2.MarginLeft = 0
    Selection.ShapeRange.TextFrame2.MarginRight = 0
    Selection.ShapeRange.TextFrame2.MarginTop = 0
    Selection.ShapeRange.TextFrame2.MarginBottom = 0
        
    Select Case �n���}�^�C�v
    Case "�`�F�b�J�[�p", "��H����", "�\��", "����[��"

        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = �� & �I���o�� & vbLf & �T�C�Y��
        '���[��n���Ȃ�1�s�ڂɃA���_�[�o�[
        If �n��s(1) = "1" Then
            Selection.ShapeRange.TextFrame2.TextRange.Characters(��a, Len(�I���o��)).Font.UnderlineStyle = msoUnderlineSingleLine
        End If
        '���[�������[�q�Ȃ�1�s�ڂ��Α�
        If �n��s(2) = "1" Then
            Selection.ShapeRange.TextFrame2.TextRange.Characters(��a, Len(�I���o��)).Font.Italic = msoTrue
        End If
        
        If Selection.Width > Selection.Height Then
            sFontSize = Selection.Height * 0.32
            gyospace = 0.6
        Else
            sFontSize = Selection.Width * 0.32
            gyospace = 0.6
        End If
        
        myFontColor = CloCodeFont1 '�t�H���g�F���x�[�X�F�Ō��߂�
        '�X�g���C�v�͌��ʂ��g��
        If clocode1 <> clocode2 Or �F�Ŕ��f = True Then
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                If �F�Ŕ��f = True Then
                    .color = 16777215 '��
                    .Radius = 12
                Else
                    .color = clocode1
                    .Radius = 8
                End If
                .color.TintAndShade = 0
                .color.Brightness = 0
                .Transparency = 0#
            End With
        End If
        
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = clofont
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 4).ParagraphFormat.SpaceWithin = 0.1 '�s��
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(Len(�I���o��) + 1, Len(�T�C�Y��) + 1).ParagraphFormat.SpaceWithin = gyospace  '�s��
    Case Else
        If Len(�T�C�Y��) > 3 Then �T�C�Y��t = Left(�T�C�Y��, 3) Else �T�C�Y��t = �T�C�Y��
        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = �� & �T�C�Y��t
        
        myFontColor = clofont '�t�H���g�F���x�[�X�F�Ō��߂�
        '�X�g���C�v�͌��ʂ��g��
        If clocode2 <> -1 Then
            With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                .color = clocode1
                .color.TintAndShade = 0
                .color.Brightness = 0
                .Transparency = 0#
                .Radius = 8
            End With
        End If
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
        'Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = &HFFFFFF And Not Selection.ShapeRange.Fill.ForeColor.RGB '�t�H���g�𔽑ΐF�ɂ���
        
        Select Case Len(�T�C�Y��t)
            Case 1
                baseSize = 1.2
            Case Else
                baseSize = 1.1
        End Select
        
        If Selection.Width > Selection.Height Then
            sFontSize = baseSize * Selection.Height * 0.4
            gyospace = 0.7
        Else
            sFontSize = baseSize * Selection.Width * 0.4
            gyospace = 0.8
        End If
    '    If Len(�|�C���g) = 4 Then
    '        sFontSize = sFontSize * 0.87
    '    End If
        
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 4).ParagraphFormat.SpaceWithin = 0.1 '�s��
        'Selection.ShapeRange.TextFrame2.TextRange.Characters(Len(�|�C���g) + 1, Len(�T�C�Y��) + 1).ParagraphFormat.SpaceWithin = gyoSpace  '�s��
    '    Stop
    End Select
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
    If �� <> "" Then
        myLen = Len(Selection.ShapeRange.TextFrame2.TextRange.Characters.Text)
        Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 2).Font.Size = sFontSize * 1
        gyospace = 0.7
        Selection.ShapeRange.TextFrame2.TextRange.Characters(1, ��a).ParagraphFormat.SpaceWithin = gyospace
    End If
        
    'line.weight��1�ɌŒ�(���ވꗗ+�̃n���}�쐬�Ɠ����ɂ���)
    Selection.ShapeRange.Line.Weight = 1
    
    �n��s = Split(�n��, "!")
    If �F�Ŕ��f = True Then
        For i2 = 1 To UBound(�n���F�ݒ�, 2)
            If �n��s(0) = �n���F�ݒ�(2, i2) Then
                Selection.Font.color = �n���F�ݒ�(1, i2)
                Exit For
            End If
        Next i2
    Else
        If �n��s(0) = "��n��" Then
            Select Case �n���\��
            Case "1"
                Call ���CH_��n��
            Case "2"
                Call ���CH_��n��_2
            Case "3"
                Call ���CH_��n��_3
            End Select
        End If
    End If
    
    If �n��s(4) <> "" Then
        If ��d�W�~flg = True Then
            If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
            Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 4
            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
        End If
    End If
    
    If �}���}1 <> "" Then
        StrName = CircleFeltTip(xLeft, yTop, myWidth, myHeight, CloCode3, clocode1, myFontColor)
        '�O���[�v��
        ActiveSheet.Shapes.Range(BaseName).Select False
        Selection.Group.Select
        Selection.Name = �[���} & "_" & cav & "_g"
    End If
    
    If �[��cav�W�� = "" Then
        �[��cav�W�� = Selection.Name
    Else
        �[��cav�W�� = �[��cav�W�� & "," & Selection.Name
    End If
    ' �쐬���ꂽ�I�[�g�V�F�C�v�̖��O��Ԃ�
    'CircleFill = Selection.Name
    'ActiveSheet.Shapes.Range(�[���}).Select False
    'Selection.Group.Select
    'ActiveSheet.Shapes.Range(Array(�[��, CircleFill)).Group.Select
    'Selection.Name = �[���}
    
End Function
Function CircleFeltTip(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, Basecolor As Long, myFontColor) As String
    
    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    
    Dim feltSize As Single
    If myWidth >= myHeight Then
        feltSize = myHeight * 0.4
    Else
        feltSize = myWidth * 0.4
    End If
    
    ' �����t���O����
    WhiteLineFrg = False
    
    ' ���~�`�`�恕�F�ݒ�
    If �n���}�^�C�v = "�`�F�b�J�[�p" Or �n���}�^�C�v = "��H����" Or �n���}�^�C�v = "�\��" Then
        feltSize = feltSize * 0.7
        ActiveSheet.Shapes.AddShape(�}���}�`��, ((xLeft + (myWidth * 1) - feltSize) * 0.747) ^ 1.0006, ((yTop + (myHeight * 1.05) - feltSize) * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006).Select
    Else
        ActiveSheet.Shapes.AddShape(�}���}�`��, ((xLeft + (myWidth * 1) - feltSize) * 0.747) ^ 1.0006, ((yTop + (myHeight * 1.05) - feltSize) * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006).Select
    End If
    Selection.ShapeRange.Glow.color.RGB = filcolor
    Selection.ShapeRange.Glow.Radius = 2
    Selection.ShapeRange.Glow.Transparency = 0.5
    Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Weight = 0.59
    
    '�}�W�b�N�F�����̎����Â炢�̂Ń��C���𔒂ɂ���
    If filcolor = 1315860 Then Selection.ShapeRange.Line.ForeColor.RGB = 16777215
    
    ' �I�[�g�V�F�C�v�̖��O��Ԃ�
    Selection.Name = �[���} & "_" & cav & "_Felt"
    CircleFeltTip = Selection.Name
    
End Function
Function BoxFeltTip(ByVal xLeft As Single, ByVal yTop As Single, ByVal myWidth As Single, ByVal myHeight As Single, filcolor As Long, Basecolor As Long, myFontColor) As String

    xLeft = xLeft * my��
    yTop = yTop * my��
    myWidth = myWidth * my��
    myHeight = myHeight * my��
    Dim feltSize As Single
        
    '�����Ɠ������O��2�ȏ㖳�����m�F_��������_�u���̃}���}
    Dim �������O�̐� As Long: �������O�̐� = 0
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        If objShp.Name = �[���} & "_" & cav & "Felt" Then
            �������O�̐� = �������O�̐� + 1
        End If
    Next
    
    If �������O�̐� > 1 Then Stop
    
    If myWidth >= myHeight Then
        feltSize = myHeight * 0.38
    Else
        feltSize = myWidth * 0.38
    End If
    ' �����t���O����
    WhiteLineFrg = False

    ' ���~�`�`�恕�F�ݒ�
    If �n���}�^�C�v = "�`�F�b�J�[�p" Or �n���}�^�C�v = "��H����" Or �n���}�^�C�v = "�\��" Then
        feltSize = feltSize * 0.7
        ActiveSheet.Shapes.AddShape(�}���}�`��, ((xLeft + (myWidth * 1) - feltSize) * 0.747) ^ 1.0006, ((yTop + (myHeight * 1.02) - feltSize) * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006).Select
    Else
        ActiveSheet.Shapes.AddShape(�}���}�`��, ((xLeft + (myWidth * 1) - feltSize) * 0.747) ^ 1.0006, ((yTop + (myHeight * 1.02) - feltSize) * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006, (feltSize * 0.747) ^ 1.0006).Select
    End If

    Selection.ShapeRange.Glow.color.RGB = filcolor
    Selection.ShapeRange.Glow.Radius = 2
    Selection.ShapeRange.Glow.Transparency = 0.5
    Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
    Selection.ShapeRange.Fill.ForeColor.RGB = filcolor
    Selection.ShapeRange.Line.Weight = 0.7

    '�}�W�b�N�F�����̎����Â炢�̂Ń��C���𔒂ɂ���
    If filcolor = 1315860 Then Selection.ShapeRange.Line.ForeColor.RGB = 16777215

    ' �I�[�g�V�F�C�v�̖��O��Ԃ�
    Selection.Name = �[���} & "_" & cav & "_Felt"
    BoxFeltTip = Selection.Name

End Function
Function ExCngColor2(TgtName As Variant) As Long
    
    
End Function

Function ExCngColorFont(TgtName As Variant) As Long
    
    ' *** �F�R�[�h�ϊ�
    
    ' �ϐ��̐錾
    Dim i As Integer                ' �ėp�ϐ�
    Dim j As Integer                ' �ėp�ϐ�
    Dim BufCloCode As Integer       ' �F�R�[�h�i�[�o�b�t�@
    Dim BufCloName As String        ' �F�L���i�[�o�b�t�@
    Dim Errfrg As Boolean           ' �G���[�t���O�ݒ�p
    
    ' �G���[�̃g���b�v
    On Error GoTo ErrSet
    
    ' �^�[�Q�b�g�𐔒l�����Ă݂�
    BufCloCode = CInt(TgtName)
    
    ' �G���[�̃g���b�v������
    On Error GoTo 0
    
    ' �G���[�͔������Ȃ�����
    If Errfrg = False Then
        
        ' ���l�͂O�`�Q�X
        If BufCloCode >= 0 And BufCloCode < 30 Then
            
            ' �ʏ�̐F�ϊ�
            ExCngColorFont = ColorVal(BufCloCode)
            
        ' ���l�͂R�O�ȏ�
        Else
            
            ' �|�C���^������
            j = 0
            
            ' ���[�v�i�R�[�h�����j
            For i = 1 To MaxRng
                
                ' ����R�[�h����������
                If ColorCode(i) = BufCloCode Then
                    
                    ' �|�C���^�擾
                    j = i
                    
                    ' ���[�v�𔲂���
                    Exit For
                    
                End If
                
            Next
            
            ' �R�[�h�ɂ��F�ϊ�
            ExCngColorFont = ColorVal(j)
            
        End If
        
    ' �G���[�����������i�����ɂ��F�w��j
    Else
        
        ' �F�L�����擾
        BufCloName = TgtName
        
        ' �S�p�����̔��p���i�S�p�����肾�Ƃ��܂��ϊ��o���Ȃ��ׁj
        ' BufCloName = WorksheetFunction.Asc(BufCloName)
        
        ' �|�C���^������
        j = 0
        
        ' ���[�v�i�L�������j
        For i = 1 To MaxRng
            
            ' ����̋L������������
            If ColorName(i) = BufCloName Then
            
                ' �|�C���^�擾
                j = i
                
                ' ���[�v�𔲂���
                Exit For
            
            End If
            
        Next
        
        ' �L���ɂ��F�ϊ�
        ExCngColorFont = ColorValFont(j)
    End If
    
    ' ���̃v���V�[�W���𔲂���
    Exit Function
    
' �G���[���̏���
ErrSet:
    
    ' �G���[�t���O�Z�b�g
    Errfrg = True
    
    ' �G���[�s�̎��̍s��
    Resume Next
    
End Function


Public Function ColorMark3(�[��, xLeft As Single, yTop As Single, myWidth As Single, myHeight As Single, �F��, ���, �`��, �}���}1, �V�[���h�t���O, �I���o��, �T�C�Y��, �n��, EmptyPlug, PlugColor, RowStr)
'    Dim �[�� As String
'    Dim xLeft As Single
'    Dim yTop As Single
'    Dim myWidth As Single
'    Dim myHeight As Single
'    Dim �F�� As String
'    Dim ��� As String
'    dim
    ' �ϐ��̐錾
    Dim i As Long               ' �ėp�ϐ�
    Dim j As Long               ' �ėp�ϐ�
    Dim LstCnt As Long          ' ���X�g���擾�p
    Dim BufMarkStr As String    ' �J���[�R�[�h�擾�p
    Dim BufClo1 As String       ' �J���[�R�[�h�P�����p
    Dim BufClo2 As String       ' �J���[�R�[�h�Q�����p
    Dim BufSize As Single       ' �T�C�Y�l�擾�p
    Dim BufRes As String        ' �`�挋�ʎ擾�p
    Dim Errfrg As Boolean       ' �G���[�t���O
    Dim lastgyo As Long
    
    'Dim myLastRow As Long: myLastRow = Cells(Rows.Count, 2).End(xlUp).Row
    'Dim ���W�͈� As Range: Set ���W�͈� = Range(Cells(41, 2), Cells(myLastRow, 9))
    'Dim ���W�͈�c As Object
    
    '�}�ȊO�̃I�u�W�F�N�g���폜
    'With Worksheets("���W")
    '    For i = .Shapes.Count To 1 Step -1
    '        If .Shapes(i).Type <> msoPicture Then
    '            If .Shapes(i).Type <> msoOLEControlObject Then
    '                .Shapes(i).Delete
    '            End If
    '        End If
    '    Next i
    'End With
        
    'For Each ���W�͈�c In ���W�͈�
    '    i = i + 1
        'Dim myProduct As String: myProduct = ���W�͈�(i, 1)
        'Dim myCav As Long: myCav = ���W�͈�(i, 2)
        'Dim xLeft As Single: xLeft = ���W�͈�(i, 3)
        'Dim yTop As Single: yTop = ���W�͈�(i, 4)
        'Dim myWidth As Single: myWidth = ���W�͈�(i, 5)
        'Dim myHeight As Single: myHeight = ���W�͈�(i, 6)
        'Dim �F�� As String: �F�ď� = ���W�͈�(i, 7)
        'Dim �`�� As String: �`�� = ���W�͈�(i, 8)
        
        'If myProduct <> myProductBack And myProductBack <> "" Then GoTo line90
        'myProductBack = myProduct
        ' �ϐ��̏�����
        'Call Init
        �n��s = Split(�n��, "!")
            ' �ݒ肳��Ă���J���[�R�[�h�擾
            BufMarkStr = �F��
            If �n����ƕ\�� <> "" And �F�� <> "" Then
                If CLng(�n��s(3)) > CLng(�n����ƕ\��) Then
                    BufMarkStr = ""
                End If
            End If
            ' �ݒ�l�����݂���
            If BufMarkStr <> "" Then
                ' �F�R�[�h�g������
                If InStr(1, BufMarkStr, "/") = 0 Then
                    BufClo1 = BufMarkStr
                    BufClo2 = "0"
                Else
                    BufClo1 = Left$(BufMarkStr, InStr(1, BufMarkStr, "/") - 1)
                    BufClo2 = Mid$(BufMarkStr, InStr(1, BufMarkStr, "/") + 1)
                End If
                If �`�� = "Box" Then BufRes = BoxFill(xLeft, yTop, myWidth, myHeight, BufClo1, GYO, BufClo2, �}���}1, CStr(�V�[���h�t���O), CStr(�I���o��), CStr(�T�C�Y��), CStr(�n��))
                If �`�� = "Cir" Then BufRes = CircleFill(GYO, xLeft, yTop, myWidth, myHeight, BufClo1, BufClo2, CStr(�}���}1), CStr(�V�[���h�t���O), CStr(�I���o��), CStr(�T�C�Y��), CStr(�n��))
                If �`�� = "Ter" Then BufRes = TerFill(xLeft, yTop, myWidth, myHeight, BufClo1, GYO, BufClo2, �}���}1, CStr(�V�[���h�t���O))
                If �`�� = "Bon" Then
                    BufRes = BonFill(xLeft, yTop, myWidth, myHeight, RowStr)
                End If
                ' ���s�ǍD
                If InStr(1, BufRes, "Err") = 0 Then
                    ' �J�E���^�C���N�������g
                    j = j + 1
                Else
                    Errfrg = True
                    Stop ' "�ϊ��ł��Ȃ��J���[�R�[�h�����݂��܂��B�n�j�{�^�����N���b�N���ďC�����Ă��������B
                    'ActiveCell.Select
                End If
            Else
                If �`�� = "Box" Then BufRes = BoxNull(xLeft, yTop, myWidth, myHeight, CStr(�I���o��), CStr(�T�C�Y��), CStr(EmptyPlug), CStr(PlugColor))
                If �`�� = "Cir" Then BufRes = CircleNull(xLeft, yTop, myWidth, myHeight, CStr(�I���o��), CStr(�T�C�Y��), CStr(EmptyPlug), CStr(PlugColor))
            End If
        If Errfrg = False Then DoEvents
    'Next ���W�͈�c

line90:
    'Set ���W�͈� = Nothing
    
End Function '

Sub �n���}�쐬_temp()

    'dictionary�p�錾
    Dim i As Long
    Dim myDic As Object, myKey, myItem
    Dim myVal, myVal2, myVal3
    Set myDic = CreateObject("Scripting.Dictionary")
    
    '�J�����g�t�H���_�ύX
    myCurDir = CurDir
    On Error Resume Next
    �ۑ��ꏊa = ActiveWorkbook.Path
    �u�b�N�� = ActiveWorkbook.Name
    CreateObject("WScript.Shell").CurrentDirectory = �ۑ��ꏊa & "\200_PVSW_RLTF"
    On Error GoTo 0
        
    '�Ώۃt�H���_���̎擾
    With Application.FileDialog(msoFileDialogFilePicker)
    If .Show = True Then �Ώۃt�@�C�� = .SelectedItems(1)
    �Ώۃt�@�C���� = Mid(�Ώۃt�@�C��, InStrRev(�Ώۃt�@�C��, "\") + 1, Len(�Ώۃt�@�C��))
    End With
    If �Ώۃt�@�C�� = "" Then End
  
    '���Ώۃt�H���_�̃t�@�C�����ɏ���
    Dim objFSO As FileSystemObject  ' FSO
    Set objFSO = New FileSystemObject
    'Set JJF = objFSO.GetFolder(�Ώۃt�@�C��).Files 'CreateObject("Scripting.FileSystemObject")
GoTo line10
    Call �œK��
        �t�@�C���p�X = �Ώۃt�@�C��
        '�t�@�C���� = JJ.Name
         
        '���e�s���ɓǂݍ���Œǉ�
            Dim strfn As String '�K�v
            strfn = �t�@�C���p�X
    
            Dim LngLoop As Long
            Dim IntFlNo As Integer
                
            IntFlNo = FreeFile
            Open strfn For Input As #IntFlNo
            
            Y = 0
                Do Until EOF(IntFlNo)
                    Y = Y + 1
                    Line Input #IntFlNo, aa
                    temp = Split(aa, ",")
                    For X = LBound(temp) To UBound(temp)
                        '���V�K�t�@�C�����쐬
                        If Y = 1 And X = 0 Then
                            �V�K�u�b�N�� = "���ވꗗ�쐬�V�X�e��_�n���}_����.xlsm"
                            �V�[�g�� = "Sheet1"
                            Workbooks.Open �ۑ��ꏊa & "\���ވꗗ�쐬�V�X�e��_�p�[�c\" & �V�K�u�b�N��
                            Dim ����() As String: ReDim ����(UBound(temp))
                            For xx = LBound(temp) To UBound(temp)
                                If Len(temp(xx)) = 15 Then ����(xx) = "0": ���i�i��count = xx + 1
                                If temp(xx) = "�n�_���L���r�e�BNo." Then ����(xx) = "0":
                                If temp(xx) = "�I�_���L���r�e�BNo." Then ����(xx) = "0"
                                If temp(xx) = "�n�_���[�����ʎq" Then ����(xx) = "0"
                                If temp(xx) = "�I�_���[�����ʎq" Then ����(xx) = "0"
                                If temp(xx) = "����" Then ����(xx) = "0"
                                With Workbooks(�V�K�u�b�N��).Sheets(�V�[�g��)
                                    .Cells(Y, xx + 1).NumberFormat = "@"
                                    .Cells(Y, xx + 1) = temp(xx)
                                End With
                            Next xx
                            Exit For
                        Else
                            With Workbooks(�V�K�u�b�N��).Sheets(�V�[�g��)
                                If ����(X) = "0" Then
                                    .Cells(Y, X + 1).NumberFormat = 0
                                Else
                                    .Cells(Y, X + 1).NumberFormat = "@"
                                End If
                                .Cells(Y, X + 1) = Replace(temp(X), vbLf, "")
                            End With
                        End If
                    Next X
                Loop
                
        Call �œK�����ǂ�
line10:
                With Workbooks(�V�K�u�b�N��).Sheets(�V�[�g��)
                    Dim �^�C�g���͈� As Range: Set �^�C�g���͈� = .Range(.Cells(1, 1), .Cells(1, .Cells(1, .Columns.count).End(xlToLeft).Column))
                    Dim �d�����ʖ�Col As Long: �d�����ʖ�Col = �^�C�g���͈�.Find("�d�����ʖ�").Column
                    Dim �d���i��Col As Long: �d���i��Col = �^�C�g���͈�.Find("�d���i��").Column
                    Dim �d���T�C�YCol As Long: �d���T�C�YCol = �^�C�g���͈�.Find("�d���T�C�Y").Column
                    Dim �d���FCol As Long: �d���FCol = �^�C�g���͈�.Find("�d���F").Column
                    Dim ����Col  As Long: ����Col = �^�C�g���͈�.Find("����").Column
                    Dim �n�_�[��Col As Long: �n�_�[��Col = �^�C�g���͈�.Find("�n�_���[�����ʎq").Column
                    Dim �I�_�[��Col As Long: �I�_�[��Col = �^�C�g���͈�.Find("�I�_���[�����ʎq").Column
                    Dim �n�_cavCol As Long: �n�_cavCol = �^�C�g���͈�.Find("�n�_���L���r�e�BNo.").Column
                    Dim �I�_cavCol As Long: �I�_cavCol = �^�C�g���͈�.Find("�I�_���L���r�e�BNo.").Column
                    Dim �n�_�[���i��Col As Long: �n�_�[���i��Col = �^�C�g���͈�.Find("�n�_���[�����i��").Column
                    Dim �I�_�[���i��Col As Long: �I�_�[���i��Col = �^�C�g���͈�.Find("�I�_���[�����i��").Column
                    Dim �n�_�}��Col As Long: �n�_�}��Col = �^�C�g���͈�.Find("�n�_���}���}�F�P").Column
                    Dim �I�_�}��Col As Long: �I�_�}��Col = �^�C�g���͈�.Find("�I�_���}���}�F�Q").Column
                    Set �^�C�g���͈� = Nothing
                End With
                Worksheets.add after:=Worksheets(Worksheets.count)
                �ǉ��V�[�g�� = ActiveSheet.Name
                lastgyo = Y
                For i = 1 To lastgyo
                    With Workbooks(�V�K�u�b�N��).Sheets(�V�[�g��)
                        Set ���i�i�Ԕ͈� = .Range(.Cells(i, 1), .Cells(i, ���i�i��count))
                        �d�����ʖ� = .Cells(i, �d�����ʖ�Col)
                        �d���i�� = .Cells(i, �d���i��Col)
                        �d���T�C�Y = .Cells(i, �d���T�C�YCol)
                        �d���F = .Cells(i, �d���FCol)
                        ���� = .Cells(i, ����Col)
                        �n�_�[�� = .Cells(i, �n�_�[��Col)
                        �I�_�[�� = .Cells(i, �I�_�[��Col)
                        �n�_cav = .Cells(i, �n�_cavCol)
                        �I�_cav = .Cells(i, �I�_cavCol)
                        �n�_�[���i�� = .Cells(i, �n�_�[���i��Col)
                        �I�_�[���i�� = .Cells(i, �I�_�[���i��Col)
                        �n�_�}�� = .Cells(i, �n�_�}��Col)
                        �I�_�}�� = .Cells(i, �I�_�}��Col)
                    End With
                    With Workbooks(�V�K�u�b�N��).Sheets(�ǉ��V�[�g��)
                        If i = 1 Then
                            .Range(.Cells(i, 1), .Cells(i, ���i�i��count)).Value = ���i�i�Ԕ͈�.Value
                            .Cells(1, ���i�i��count + 1) = "�d�����ʖ�": .Columns(���i�i��count + 1).NumberFormat = "@"
                            .Cells(1, ���i�i��count + 2) = "�d���i��": .Columns(���i�i��count + 2).NumberFormat = "@"
                            .Cells(1, ���i�i��count + 3) = "�d���T�C�Y": .Columns(���i�i��count + 3).NumberFormat = "@"
                            .Cells(1, ���i�i��count + 4) = "�d���F": .Columns(���i�i��count + 4).NumberFormat = "@"
                            .Cells(1, ���i�i��count + 5) = "����": .Columns(���i�i��count + 5).NumberFormat = 0
                            .Cells(1, ���i�i��count + 6) = "�[����": .Columns(���i�i��count + 6).NumberFormat = 0
                            .Cells(1, ���i�i��count + 7) = "cav": .Columns(���i�i��count + 7).NumberFormat = 0
                            .Cells(1, ���i�i��count + 8) = "�}���}": .Columns(���i�i��count + 8).NumberFormat = "@"
                            .Cells(1, ���i�i��count + 9) = "���i�i��": .Columns(���i�i��count + 9).NumberFormat = "@"
                        Else
                            addgyo = .Cells(.Rows.count, ���i�i��count + 1).End(xlUp).Row + 1
                            .Range(.Cells(addgyo, 1), .Cells(addgyo + 1, ���i�i��count)).Value = ���i�i�Ԕ͈�.Value
                            .Cells(addgyo, ���i�i��count + 1) = �d�����ʖ�
                            .Cells(addgyo, ���i�i��count + 2) = �d���i��
                            .Cells(addgyo, ���i�i��count + 3) = �d���T�C�Y
                            .Cells(addgyo, ���i�i��count + 4) = �d���F
                            .Cells(addgyo, ���i�i��count + 5) = ����
                            .Cells(addgyo, ���i�i��count + 6) = �n�_�[��
                            .Cells(addgyo, ���i�i��count + 7) = �n�_cav
                            .Cells(addgyo, ���i�i��count + 8) = �n�_�}��
                            .Cells(addgyo, ���i�i��count + 9) = �n�_�[���i��
                            .Cells(addgyo + 1, ���i�i��count + 1) = �d�����ʖ�
                            .Cells(addgyo + 1, ���i�i��count + 2) = �d���i��
                            .Cells(addgyo + 1, ���i�i��count + 3) = �d���T�C�Y
                            .Cells(addgyo + 1, ���i�i��count + 4) = �d���F
                            .Cells(addgyo + 1, ���i�i��count + 5) = ����
                            .Cells(addgyo + 1, ���i�i��count + 6) = �I�_�[��
                            .Cells(addgyo + 1, ���i�i��count + 7) = �I�_cav
                            .Cells(addgyo + 1, ���i�i��count + 8) = �I�_�}��
                            .Cells(addgyo + 1, ���i�i��count + 9) = �I�_�[���i��
                        End If
                    End With
                Next i
            Stop
                With Workbooks(�V�K�u�b�N��).Sheets(�ǉ��V�[�g��)
                    .Range(.Cells(1, 1), .Cells(addgyo + 1, ���i�i��count + 9)).Sort _
                        key1:=Range("j2"), Order1:=xlAscending, _
                        key2:=Range("k2"), order2:=xlAscending, _
                        Header:=xlYes
                    lastgyo = .Cells(.Rows.count, ���i�i��count + 1).End(xlUp).Row
                End With
              Stop
                With Workbooks(�u�b�N��).Sheets("���W")
                    Set ���W�͈� = .Range(.Cells(41, 2), .Cells(.Cells(.Rows.count, 2).End(xlUp).Row, 11))
                    ���Wlastgyo = .Range("b" & .Rows.count).End(xlUp).Row
                End With
                For i = 2 To lastgyo
                    With Workbooks(�V�K�u�b�N��).Sheets(�ǉ��V�[�g��)
                        �d���T�C�Y = .Cells(i, ���i�i��count + 3)
                        �d���F = .Cells(i, ���i�i��count + 4)
                        �[�� = .Cells(i, ���i�i��count + 6)
                        cav = .Cells(i, ���i�i��count + 7)
                        �}���} = .Cells(i, ���i�i��count + 8)
                        ���i�i�� = .Cells(i, ���i�i��count + 9)
                    End With
                    With Workbooks(�u�b�N��).Sheets("���W")
                        For j = 1 To ���Wlastgyo
                            If Replace(���W�͈�(j, 1), "-", "") = ���i�i�� Then Stop
                        Next
                        If ���i�i�� <> ���i�i��back And ���i�i��back <> "" Then
                            '�}�̓ǂݍ���
                            URL = �ۑ��ꏊa & "\���ވꗗ�쐬�V�X�e��_���}\" & ���i�i�� & "_0_00"
                            ActiveSheet.Pictures.Insert("D:\18_���ވꗗ\���ވꗗ�쐬�V�X�e��_���}\7009-1323_0_001.emf").Select
                        End If
                        ���i�i��back = ���i�i��
                    End With
                Next i
              
                                �i�� = Replace(Mid(temp(X), 1, 15), " ", "")
                                �ݕ� = Mid(temp(X), 19, 3)
                            '�i�Ԃ��قȂ�ꍇ
                            If �i�� <> �i��bak Then GoSub ���i���X�g����
                            With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��) '�^�C�g��
                                    lastgyo = .Range("a" & .Rows.count).End(xlUp).Row + 1
                                If lastgyo = 2 And .Cells(1, 1) = "" Then
                                    .Range(.Cells(1, 1), .Cells(1, .Columns.count)).NumberFormat = "@"
                                    .Cells(1, 1) = "���i�i��"
                                    .Cells(1, 2) = "�ݕ�"
                                    .Cells(1, 3) = "���i�i��"
                                    .Cells(1, 4) = "-"
                                    .Cells(1, 5) = "����"
                                    .Cells(1, 7) = "�F"
                                    .Cells(1, 8) = "����"
                                    .Cells(1, 9) = "�ӏ���"
                                    .Cells(1, 10) = "��ƍH��"
                                    .Cells(1, 11) = "˻��"
                                    .Cells(1, 13) = "�ď�"
                                End If
                                '����ȕ���,�`���[�u��
                                If Mid(temp(X), 27, 1) = "T" Then
                                    .Cells(lastgyo, 1) = Mid(temp(X), 1, 15)
                                    .Cells(lastgyo, 2) = Mid(temp(X), 19, 3)
                                    .Cells(lastgyo, 3) = Mid(temp(X), 375, 8)
                                    .Cells(lastgyo, 4) = Replace(Mid(temp(X), 383, 6), " ", "")
                                    .Cells(lastgyo, 5) = Replace(Mid(temp(X), 389, 4), " ", "")
                                    .Cells(lastgyo, 6) = Replace(Mid(temp(X), 393, 4), " ", "")
                                    If .Cells(lastgyo, 5) = "" And .Cells(lastgyo, 6) = "" Then
                                        �T�C�Y = .Cells(lastgyo, 5) & .Cells(lastgyo, 6)
                                    ElseIf .Cells(lastgyo, 5) <> "" And .Cells(lastgyo, 6) = "" Then
                                        �T�C�Y = "D" & String(3 - Len(.Cells(lastgyo, 5)), " ") & .Cells(lastgyo, 5)
                                    ElseIf .Cells(lastgyo, 5) = "" And .Cells(lastgyo, 6) <> "" Then
                                        �T�C�Y = .Cells(lastgyo, 5)
                                    Else
                                        aaaa = String(3 - Len(Replace(.Cells(lastgyo, 5), ".", "")), " ") & .Cells(lastgyo, 5)
                                        aaaB = String(3 - Len(Replace(.Cells(lastgyo, 6), ".", "")), " ") & .Cells(lastgyo, 6)
                                        �T�C�Y = aaaa & "�~" & aaaB
                                    End If
                                    .Cells(lastgyo, 7) = Mid(temp(X), 397, 5)
                                    .Cells(lastgyo, 8) = Mid(temp(X), 403, 5)
                                        ���� = "L=" & String(4 - Len(.Cells(lastgyo, 8)), " ") & .Cells(lastgyo, 8)
                                    .Cells(lastgyo, 9) = 1
                                    .Cells(lastgyo, 10) = Mid(temp(X), 153, 2)
                                    .Cells(lastgyo, 11).NumberFormat = "@"
                                    .Cells(lastgyo, 11) = Left(Mid(temp(X), 544, 4), 1)
                                    .Cells(lastgyo, 12) = Mid(temp(X), 544, 4)
                                    .Cells(lastgyo, 13) = Left(.Cells(lastgyo, 3), 3) & "-" & �T�C�Y & " " & ����
                                '�R�l�N�^�ނ̕���
                                ElseIf Mid(temp(X), 27, 1) = "B" Then
                                    For bs = 0 To 9
                                        If Mid(temp(X), 174 + (bs * 20) + 11, 3) = "ATO" Then
                                            lastgyo = .Range("a" & .Rows.count).End(xlUp).Row + 1
                                            .Cells(lastgyo, 1) = Mid(temp(X), 1, 15)
                                            .Cells(lastgyo, 2) = Mid(temp(X), 19, 3)
                                            .Cells(lastgyo, 3) = Replace(Mid(temp(X), 175 + (bs * 20), 10), " ", "")
                                            .Cells(lastgyo, 4) = " "
                                            .Cells(lastgyo, 9) = Val(Mid(temp(X), 175 + (bs * 20) + 15, 4))
                                            .Cells(lastgyo, 10) = Mid(temp(X), 558 + (bs * 2), 2)
                                            .Cells(lastgyo, 11).NumberFormat = "@"
                                            .Cells(lastgyo, 11) = Left(Mid(temp(X), 544, 4), 1)
                                            .Cells(lastgyo, 12) = Mid(temp(X), 544, 4)
                                            Select Case Len(.Cells(lastgyo, 3))
                                            Case 8
                                            .Cells(lastgyo, 13) = Mid(.Cells(lastgyo, 3), 1, 4) & "-" & Mid(.Cells(lastgyo, 3), 5, 4)
                                            Case 10
                                            .Cells(lastgyo, 13) = Mid(.Cells(lastgyo, 3), 1, 4) & "-" & Mid(.Cells(lastgyo, 3), 5, 4) & "-" & Mid(.Cells(lastgyo, 3), 9, 2)
                                            Case Else
                                            .Cells(lastgyo, 13) = .Cells(lastgyo, 3)
                                            End Select
                                        End If
                                    Next bs
                                End If
                            End With
                            If EOF(IntFlNo) Then GoSub ���i���X�g����
                                '���V�K�t�@�C�����쐬
                                If Y = 1 And X = 0 Then
                                    Workbooks.add
                                    �V�K�u�b�N�� = ActiveWorkbook.Name
                                    �V�K�V�[�g�� = ActiveSheet.Name
                                End If
                                
                                With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��)
                                '���ŏ��̐ݒ�
                                If Y = ��������s Then
                                        .Range(.Cells(1, 1), .Cells(1, .Columns.count)).NumberFormat = �����s����
                                        '�������̐ݒ�
                                        '�����Ɋ܂܂�镶���񂪂���ꍇ�͏���=0
                                        Select Case True
                                            Case ���� Like "*_" & temp(X) & "_*"
                                            If X < 256 Then .Range(.Cells(Y + 1, X + 1), .Cells(.Rows.count, X + 1)).NumberFormat = 0
                                            Case ���� Like "*_" & temp(X) & "_*"
                                            If X < 256 Then .Range(.Cells(Y + 1, X + 1), .Cells(.Rows.count, X + 1)).NumberFormat = ���ꏑ��
                                            Case Else
                                                If X < 256 Then
                                                .Range(.Cells(Y + 1, X + 1), .Cells(.Rows.count, X + 1)).NumberFormat = ���̑�����
                                                End If
                                        End Select
                                        'LEN15�̏ꍇ�͏���=0
                                        If Len(temp(X)) = 15 And X < 256 Then
                                            .Range(.Cells(Y + 1, X + 1), .Cells(.Rows.count, X + 1)).NumberFormat = 0
                                            �i�Ԉꗗ = �i�Ԉꗗ & Replace(temp(X), " ", "") & vbLf
                                        End If
                                End If
                                    '���l���o��
                                    If X < 256 Then .Cells(Y, X + 1) = temp(X)
                                End With
            If �^�C�g�� <> "���i���X�g" Then
                Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��).Columns.AutoFit
                '���ۑ�
                 �ۑ��ꏊ = �ۑ��ꏊa & "\002_�G�N�Z���f�[�^\" & �Ώۃt�H���_�� & "\" & �Ώۃt�H���_�� & "_" & �^�C�g�� & ".xls"
                 If �^�C�g�� = "���i�ʉ�H�}�g���N�X" Then �t�@�C����p = �ۑ��ꏊ
                 If Dir(�ۑ��ꏊa & "\002_�G�N�Z���f�[�^\" & �Ώۃt�H���_��, vbDirectory) = "" Then MkDir (�ۑ��ꏊa & "\002_�G�N�Z���f�[�^\" & �Ώۃt�H���_��)
                 Application.DisplayAlerts = False
                 ActiveWorkbook.SaveAs fileName:=�ۑ��ꏊ, FileFormat:=xlExcel8
                 Application.DisplayAlerts = True
                 ActiveWorkbook.Close
             End If
        Close #IntFlNo
        
line20:

    Set objFSO = Nothing
    Set JJF = Nothing
  
    Cells(5, 3) = �i�Ԉꗗ
    Call �œK�����ǂ�
Exit Sub

���i���X�g����:
With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��)
    '�����בւ�
    .Range("a2:z" & .Range("a" & .Rows.count).End(xlUp).Row).Sort _
    key1:=.Range("m2"), Order1:=xlAscending, _
    Header:=xlGuess
    .Range("a2:z" & .Range("a" & .Rows.count).End(xlUp).Row).Sort _
    key1:=.Range("j2"), Order1:=xlAscending, _
    key2:=.Range("k2"), order2:=xlAscending, _
    key3:=.Range("d2"), Order3:=xlAscending, _
    Header:=xlGuess
    .Columns.AutoFit
End With
        '���H��50
        With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��)
            ���i�i�� = .Range("a2") & "_" & .Range("b2")
            '�����f�[�^��z��Ɋi�[
            myVal = .Range("A2", .Range("A" & .Rows.count).End(xlUp)).Resize(, 20).Value
                'myDic�փf�[�^���i�[
                For i = 1 To UBound(myVal, 1)
                    If myVal(i, 10) = "50" Or myVal(i, 10) = "60" Or myVal(i, 10) = "70" Or myVal(i, 10) = "80" Then
                        myVal2 = myVal(i, 13) & "_" & myVal(i, 10)
                        If Not myDic.exists(myVal2) Then
                            myDic.add myVal2, myVal(i, 9)
                        Else
                            myDic(myVal2) = myDic(myVal2) + myVal(i, 9)
                        End If
                    End If
                Next i
        End With
        
        Workbooks(�V�K�u�b�N��).Sheets("����A3").Copy Workbooks(�V�K�u�b�N��).Sheets("����A3")
        Workbooks(�V�K�u�b�N��).ActiveSheet.Name = "Auto(�T�u��)"
        �V�K�V�[�g��2 = ActiveSheet.Name
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2).Range("s4") = "���i�i��:    " & �i��bak & "  " & �ݕ� & "     "
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2).Range("ak1") = "�� �� ���F'" & Right(Year(Date), 2) & "�N " & Format((Date), "mm") & "�� " & Format(Date, "dd") & "��"
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2).Range("ae5") = "���i�ԕʃ��X�g"
        '��Key,Item�̏����o��
        With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2)
        myKey = myDic.keys
        myItem = myDic.items
            lastcolumn = 19: lastRow = 8: co = 0
            For i = 0 To UBound(myKey)
                myVal3 = Split(myKey(i), "_")
                '50���I������玟�̗�
                If ��ƍH��bak = "50" And ��ƍH��bak <> myVal3(1) And co <> 0 Then lastcolumn = lastcolumn + 6: lastRow = 8: co = 0
                .Cells(lastRow + co, lastcolumn).Value = myVal3(0)
                .Cells(lastRow + co, lastcolumn + 1).Value = myItem(i)
                Select Case myVal3(1)
                Case "50": cc = 2
                Case "60": cc = 3
                Case "70": cc = 4
                Case "80": cc = 5
                Case Else: Stop  '��ƍH������L�ȊO
                End Select
                .Cells(lastRow + co, lastcolumn + cc).Value = "��"
                co = co + 1
                If co = 30 Then lastcolumn = lastcolumn + 6: lastRow = 8: co = 0
                ��ƍH��bak = myVal3(1)
            Next i
        End With
        'dictionary���ăZ�b�g
         Set myDic = Nothing
         Set myDic = CreateObject("Scripting.Dictionary")
                
        '���H��40_�T�u��
        With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��)
            '�����f�[�^��z��Ɋi�[
            myVal = .Range("A2", .Range("A" & .Rows.count).End(xlUp)).Resize(, 20).Value
                'myDic�փf�[�^���i�[
                For i = 1 To UBound(myVal, 1)
                    If myVal(i, 10) = "40" Then
                        myVal2 = myVal(i, 13) & "_" & myVal(i, 12) & "_" & myVal(i, 10)
                        If Not myDic.exists(myVal2) Then
                            myDic.add myVal2, myVal(i, 9)
                        Else
                            myDic(myVal2) = myDic(myVal2) + myVal(i, 9)
                        End If
                    End If
                Next i
        End With
        
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2).Range("b4") = "���i�i��:    " & �i��bak & "  " & �ݕ� & "     "
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2).Range("n1") = "�� �� ���F'" & Right(Year(Date), 2) & "�N " & Format((Date), "mm") & "�� " & Format(Date, "dd") & "��"
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2).Range("j6") = "���T�u�ʃ��X�g"
        '��Key,Item�̏����o��
        With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2)
        myKey = myDic.keys
        myItem = myDic.items
            lastcolumn = 2: lastRow = 8: co = 0: cc = 0
            For i = 0 To UBound(myKey)
                myVal3 = Split(myKey(i), "_")
                '�T�uL1���O��ƈقȂ�ꍇ�A1�s�󂯂�
                If �T�uL1bak <> Left(myVal3(1), 1) And co <> 0 Then co = co + 1
                '1�s�󂯂�30�s�𒴂����ꍇ�͎��̗�
                If co = 30 And lastcolumn <> 14 Then lastcolumn = lastcolumn + 3: lastRow = 8: co = 0
                .Cells(lastRow + co, lastcolumn).Value = myVal3(0)
                .Cells(lastRow + co, lastcolumn + 1).Value = myItem(i)
                .Cells(lastRow + co, lastcolumn + 2).Value = Replace(myVal3(1), " ", "")
                co = co + 1
                If co = 30 And lastcolumn <> 14 Then lastcolumn = lastcolumn + 3: lastRow = 8: co = 0
                �T�uL1bak = Left(myVal3(1), 1)
            Next i
        End With
        'dictionary���ăZ�b�g
         Set myDic = Nothing
         Set myDic = CreateObject("Scripting.Dictionary")
         
        '���H��40_�i�ԏ�
        With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��)
            '�����בւ�
            .Range("a2:z" & .Range("a" & .Rows.count).End(xlUp).Row).Sort _
            key1:=.Range("d2"), Order1:=xlAscending, _
            key2:=.Range("m2"), order2:=xlAscending, _
            key3:=.Range("l2"), Order3:=xlAscending, _
            Header:=xlGuess
            .Columns.AutoFit
        End With
        
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2).Copy Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2)
        ActiveSheet.Name = "Auto(�i�ԕ�)"
        �V�K�V�[�g��3 = ActiveSheet.Name
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2).Columns("s:ap").Delete
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��3).Range("b8:p" & Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��2).Rows.count).ClearContents
        Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��3).Range("j6") = "���i�ԕʃ��X�g"
        With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��)
            '�����f�[�^��z��Ɋi�[
            myVal = .Range("A2", .Range("A" & .Rows.count).End(xlUp)).Resize(, 20).Value
                'myDic�փf�[�^���i�[
                For i = 1 To UBound(myVal, 1)
                    If myVal(i, 10) = "40" Then
                        myVal2 = myVal(i, 13) & "_" & myVal(i, 12) & "_" & myVal(i, 10)
                        If Not myDic.exists(myVal2) Then
                            myDic.add myVal2, myVal(i, 9)
                        Else
                            myDic(myVal2) = myDic(myVal2) + myVal(i, 9)
                        End If
                    End If
                Next i
        End With
        
        '��Key,Item�̏����o��
        With Workbooks(�V�K�u�b�N��).Sheets(�V�K�V�[�g��3)
        myKey = myDic.keys
        myItem = myDic.items
            lastcolumn = 2: lastRow = 7: co = 0: cc = 0
            For i = 0 To UBound(myKey)
                myVal3 = Split(myKey(i), "_")
                '�T�uL1���O��ƈقȂ�ꍇ�A1�s�󂯂�
                If �ď�L4bak <> Left(myVal3(0), 4) And co <> 0 Then co = co + 1
                '1�s�󂯂�30�s�𒴂����ꍇ�͎��̗�
                If co >= 30 And lastcolumn <> 14 Then lastcolumn = lastcolumn + 3: lastRow = 7: co = 0
                '�ď̂�4�����ڂ�-�̏ꍇ�͎��̗�
                If �ď�M4_1 <> Mid(myVal3(0), 4, 1) And Mid(myVal3(0), 4, 1) = "-" And lastcolumn <> 14 And co <> 0 Then lastcolumn = lastcolumn + 3: lastRow = 7: co = 0
                '���i�i�Ԃ������ꍇ�A�����s�ɂ܂Ƃ߂�
                If �ď�bak = myVal3(0) And co <> 0 Then
                .Cells(lastRow + co, lastcolumn + 1) = .Cells(lastRow + co, lastcolumn + 1).Value + myItem(i)
                .Cells(lastRow + co, lastcolumn + 2) = .Cells(lastRow + co, lastcolumn + 2) & "�" & Replace(myVal3(1), " ", "")
                Else
                co = co + 1
                .Cells(lastRow + co, lastcolumn).Value = myVal3(0)
                .Cells(lastRow + co, lastcolumn + 1).Value = .Cells(lastRow + co, lastcolumn + 1).Value + myItem(i)
                .Cells(lastRow + co, lastcolumn + 2).Value = Replace(myVal3(1), " ", "")
                End If
                If co = 30 And lastcolumn <> 14 Then lastcolumn = lastcolumn + 3: lastRow = 7: co = 0
                �ď�L4bak = Left(myVal3(0), 4)
                �ď�bak = myVal3(0)
                �ď�M4_1 = Mid(myVal3(0), 4, 1)
            Next i
        End With
        'dictionary���ăZ�b�g
         Set myDic = Nothing
         Set myDic = CreateObject("Scripting.Dictionary")
        
        Application.DisplayAlerts = False
        Workbooks(�V�K�u�b�N��).Sheets("����A3").Delete
        Application.DisplayAlerts = True
            '���ۑ�
             �ۑ��ꏊ = �ۑ��ꏊa & "\002_�G�N�Z���f�[�^\" & �Ώۃt�H���_�� & "\���i���X�g\" & �Ώۃt�H���_�� & "_" & �^�C�g�� & "_" & �i��bak & "_" & �ݕ�bak & ".xls"
             If �^�C�g�� = "���i�ʉ�H�}�g���N�X" Then �t�@�C����p = �ۑ��ꏊ
             If Dir(�ۑ��ꏊa & "\002_�G�N�Z���f�[�^\" & �Ώۃt�H���_��, vbDirectory) = "" Then MkDir (�ۑ��ꏊa & "\002_�G�N�Z���f�[�^\" & �Ώۃt�H���_��)
             If Dir(�ۑ��ꏊa & "\002_�G�N�Z���f�[�^\" & �Ώۃt�H���_�� & "\���i���X�g", vbDirectory) = "" Then MkDir (�ۑ��ꏊa & "\002_�G�N�Z���f�[�^\" & �Ώۃt�H���_�� & "\���i���X�g")
             
             Application.DisplayAlerts = False
             ActiveWorkbook.SaveAs fileName:=�ۑ��ꏊ, FileFormat:=xlExcel8
             Application.DisplayAlerts = True
             ActiveWorkbook.Close
             
             �i��bak = �i��
             �ݕ�bak = �ݕ�
            
            '���V�K�t�@�C�����쐬
            If EOF(IntFlNo) = False Then
                �V�K�u�b�N�� = "�P�����͕ϊ��V�X�e��_����.xls"
                �V�K�V�[�g�� = "�ꗗ"
                Workbooks.Open ActiveWorkbook.Path & "\000_����\" & �V�K�u�b�N��
                'dictionary���ăZ�b�g
                Set myDic = Nothing
                Set myDic = CreateObject("Scripting.Dictionary")
            End If
Return


End Sub
Function ���CH_Ver182_MoreLater()
    
    'Debug.Print Application.Caller
    'ActiveSheet.Shapes.Range(Application.Caller).Select
    Application.StatusBar = Application.Caller
    'ActiveSheet.Shapes.Range(Application.Caller).Select
    If ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor = RGB(255, 80, 80) Then
        ActiveSheet.Shapes.Range(Application.Caller).Line.Weight = ActiveSheet.Shapes.Range(Application.Caller).Line.Weight / 6
        ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(20, 20, 20)
        Dim filcolor As Long
        filcolor = ActiveSheet.Shapes.Range(Application.Caller).Fill.ForeColor
        If filcolor = 1315860 Then ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
        ActiveSheet.Shapes.Range(Application.Caller).Line.Weight = ActiveSheet.Shapes.Range(Application.Caller).Line.Weight * 6
        ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(255, 80, 80)
        'Selection.Name = Selection.Name & "s"
    End If
    'Selection.Unselect
    'Cells(1, 1).Select
    '�x�[�X�F�����̏ꍇ�A���ɂ���
        
    'select����������ׂɎʐ^��I��remake
    Dim myShapeName As String
    Dim a, b, i As Long
    myShapeName = Application.Caller
    For i = 1 To Len(myShapeName)
        If Mid(myShapeName, i, 1) = "_" Then a = a + 1
        If a = 2 Then
            b = i
            Exit For
        End If
    Next i
    myShapeName = Left(myShapeName, b - 1)
    ActiveSheet.Shapes.Range(myShapeName).Select
    
End Function
Function ���CH()
    
    'Debug.Print Application.Caller
    'ActiveSheet.Shapes.Range(Application.Caller).Select
    'Application.StatusBar = Application.Caller
    'ActiveSheet.Shapes.Range(Application.Caller).Select
    If ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor = RGB(255, 80, 80) Then
        ActiveSheet.Shapes.Range(Application.Caller).Line.Weight = ActiveSheet.Shapes.Range(Application.Caller).Line.Weight / 5
        ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(20, 20, 20)
        Dim filcolor As Long
        filcolor = ActiveSheet.Shapes.Range(Application.Caller).Fill.ForeColor
        If filcolor = 1315860 Then ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
        ActiveSheet.Shapes.Range(Application.Caller).Line.Weight = ActiveSheet.Shapes.Range(Application.Caller).Line.Weight * 5
        ActiveSheet.Shapes.Range(Application.Caller).Line.ForeColor.RGB = RGB(255, 80, 80)
        'Selection.Name = Selection.Name & "s"
    End If
    'Selection.Unselect
    'Cells(1, 1).Select
    '�x�[�X�F�����̏ꍇ�A���ɂ���
        
    'select����������ׂɎʐ^��I��remake
    Dim myShapeName As String
    Dim a, b, i As Long
    myShapeName = Application.Caller
    For i = 1 To Len(myShapeName)
        If Mid(myShapeName, i, 1) = "_" Then a = a + 1
        If a = 2 Then
            b = i
            Exit For
        End If
    Next i
    myShapeName = Left(myShapeName, b - 1)
    ActiveSheet.Shapes.Range(myShapeName).Select
    
End Function
Function ���CH_��n��()
    
    'Selection.ShapeRange.Line.DashStyle = 1
'    'Selection.ShapeRange.Line.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
'    'Selection.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line.Weight = 8
'    Selection.ShapeRange.Adjustments.Item(1) = 0.15
'    'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
'    Selection.ShapeRange.Line.OneColorGradient msoGradientDiagonalUp, 1, 1
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.4
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.401
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.599
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.6
'    Selection.ShapeRange.Line.GradientStops.Insert Filcolor, 0.99
'    Selection.ShapeRange.Line.GradientStops.Delete 1
'    Selection.ShapeRange.Line.GradientStops.Delete 1


    Selection.ShapeRange.Line.ForeColor.RGB = RGB(0, 250, 250)

    If Selection.ShapeRange.Line.Weight < 0 Then Selection.ShapeRange.Line.Weight = 0.1
    Selection.ShapeRange.Line.Weight = Selection.ShapeRange.Line.Weight * 3
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 80, 80)
    
End Function


Function ���CH_��n��_2()
    With Selection
        .Left = .Left + .Width * 0.25
        .Top = .Top + .Height * 0.25
        .Width = .Width * 0.5
        .Height = .Height * 0.5
        '.Line.Weight = .Line.Weight
        .ShapeRange.TextFrame2.TextRange.Characters.Text = ""
    End With
End Function
Function ���CH_��n��_3()
    With Selection.ShapeRange
        .Fill.Patterned msoPatternWideDownwardDiagonal
        .Fill.ForeColor.RGB = RGB(255, 255, 250)
        .Fill.BackColor.RGB = RGB(20, 20, 20)
        .TextFrame2.TextRange.Characters.Text = ""
        .Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
End Function
Function ���Make(Optional �摜�� As String)
    If �摜�� = "" Then �摜�� = Application.Caller
    With ActiveWorkbook.Sheets("���i�i��")
        Dim ��n���}�\�� As Long: ��n���}�\�� = .Range("s4")
        If Not (��n���}�\�� = 1 Or ��n���}�\�� = 2) Then MsgBox "Sheets(���i�i��)��Cells(S4)��1��2�őI�����Ă��������B": End
    End With
    
    Call ���`��_SH_season2(�摜��)
    Call ���`��_AH(��n���}�\��, �摜��)
'Exit Function
    'If ActiveSheet.Name = "PVSW_�n���}" Then Call ���CH_season2
    
End Function

Function ���CH_season2()
    'Debug.Print Application.Caller
    Call �œK��
    
    With ActiveWorkbook.Sheets("PVSW_�n���}")
        Dim ���i�i�ԓ_�� As Long: ���i�i�ԓ_�� = .Cells.Find("�[�����i��").Column - 1
    End With
    
    Dim a, b, c, z, i As Long
    Dim yoso
    Dim �[���} As String: �[���} = Left(Application.Caller, Len(Application.Caller) - 2)
    Dim �[���}Len As Long: �[���}Len = Len(�[���})
    Dim objShp As Shape
        For Each objShp In ActiveSheet.Shapes
            If �[���} = Left(objShp.Name, �[���}Len) Then
                a = 0: b = 0: c = 0: z = 0
                For i = 1 To Len(objShp.Name)
                    If Mid(objShp.Name, i, 1) = "_" Then z = z + 1
                    If z = 1 And a = 0 Then a = i
                    If z = 2 And b = 0 Then b = i
                    If z = 3 And c = 0 Then c = i
                Next i
                If InStr(objShp.Name, "w") > 0 Then
                    myCav = Mid(objShp.Name, b + 1, c - b - 1)
                Else
                    myCav = Mid(objShp.Name, b + 1, Len(objShp.Name) - b)
                End If
                If IsNumeric(myCav) Then
                    yoso = Split(objShp.Name, "_")
                    Dim �[�� As String: �[�� = yoso(0)
                    Dim �g���� As String: �g���� = yoso(1)
                    Dim cav As String: cav = yoso(2)
                    Dim �n���� As String
                    If ActiveSheet.Shapes.Range(objShp.Name).Line.ForeColor.RGB = RGB(255, 80, 80) Then
                        �n���� = "��"
                    Else
                        �n���� = "��"
                    End If
                    'sheets("PVSW_RLTF���[")�ւ̔��f
                    With Sheets("PVSW_RLTF���[")
                        Dim �d�����ʖ�Col As Long: �d�����ʖ�Col = .Cells.Find("�d�����ʖ�").Column
                        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 5).End(xlUp).Row
                        Dim �[��Col As Long: �[��Col = .Cells.Find("�[�����ʎq").Column
                        Dim cavCol As Long: cavCol = .Cells.Find("�L���r�e�BNo.").Column
                        Dim �n����Col As Long: �n����Col = .Cells.Find("�n����").Column
                        Dim X As Long
                        
                        Dim varBinary As Variant: varBinary = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", _
                        "1000", "1001", "1010", "1011", "1100", "1101", "1110", "1111")
                        Dim strH As String: strH = �g����
                        Dim HtoB As String
                        ReDim strhtob(1 To Len(strH)) As String
                        For i = 1 To Len(strH)
                            strhtob(i) = varBinary(Val("&H" & Mid$(strH, i, 1)))
                        Next i
                        HtoB = Join$(strhtob, vbNullString)
                        �g���� = Right(HtoB, 10)
                        
                        For X = 1 To Len(�g����)
                            If Mid(�g����, X, 1) = 1 Then
                                For i = 2 To lastRow
                                    If .Cells(i, �d�����ʖ�Col) <> "" Then
                                        If .Cells(i, �[��Col) = �[�� Then
                                            If .Cells(i, cavCol) = cav Then
                                                .Cells(i, �n����Col) = �n����
                                            End If
                                        End If
                                    End If
                                Next i
                            End If
                        Next X
                    End With
                    '���̃V�[�g�ւ̔��f
                    With ActiveSheet
                        Dim �\��Col As Long: �\��Col = .Cells.Find("�\��").Column
                        Dim �F As Long
                        �[��Col = .Cells.Find("�[����").Column
                        cavCol = .Cells.Find("Cav").Column
                        lastRow = .Cells(.Rows.count, �\��Col).End(xlUp).Row
                        For i = 2 To lastRow
                            If �[�� = .Cells(i, �[��Col) Then
                                If cav = .Cells(i, cavCol) Then
                                    For X = 1 To Len(Right(�g����, ���i�i�ԓ_��))
                                        If .Cells(i, X) <> "" Then
                                            If Mid(Right(�g����, ���i�i�ԓ_��), X, 1) = Val(Replace(.Cells(i, X), " ", "")) Then
                                                If �n���� = "��" Then
                                                    �F = RGB(240, 150, 150)
                                                Else
                                                    �F = RGB(150, 150, 240)
                                                End If
                                                .Cells(i, X).Interior.color = �F
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next i
                    End With
                End If
                'objShp.Duplicate
                'Selection.Name = objShp.Name & "s"
            End If
        Next
        Call �œK�����ǂ�
End Function

Public Function ���`��_AH(��n���}�\�� As Long, Optional �摜�� As String)
Dim pTime As Single: pTime = Timer
    Dim �[���} As String: �[���} = Left(�摜��, Len(�摜��) - 2)
    Dim �[���}Len As Long: �[���}Len = Len(�[���})
    Dim myLeft, myTop, myWidth, myHeight As Long
    Dim myLeftM As Long
    Dim myTopM As Long
    Dim myWidthM As Long
    Dim myHieghtM As Long
    Dim myLineWeight As Long
    Dim myFlag As Long
    Dim myFlagName As String
    Dim objShp As Shape
    Dim aa As Boolean
    Dim myFlagRange As Variant
    Dim c As Variant
    Dim skipFlag As Long
    Dim fontDel As Long
    On Error Resume Next
    aa = IsObject(ActiveSheet.Shapes.Range(�[���} & "_AH"))
    On Error GoTo 0
    If aa = True Then ActiveSheet.Shapes.Range(�[���} & "_AH").Delete
        
        For Each objShp In ActiveSheet.Shapes
            If �[���} = Left(objShp.Name, �[���}Len) And Not objShp.Name Like "*SH*" Then
            'Debug.Print objShp.Name
                skipFlag = 0
                If myFlagName <> "" Then
                    myFlagRange = Split(myFlagName, ",")
                    For Each c In myFlagRange
                        If c <> "" Then
                            If Left(objShp.Name, InStrRev(objShp.Name, "_")) & "e" = c Then
                                skipFlag = 1
                                Exit For
                            End If
                        End If
                    Next c
                    'kkk = InStr(myFlagName, objShp.Name)
                End If
                If skipFlag = 1 And ��n���}�\�� = 2 Then GoTo line20
                
                If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                    myFlagName = myFlagName & "," & objShp.Name & "_e"
                End If
                
                fontDel = 0
                If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                    Select Case ��n���}�\��
                    Case 1
                        myLeft = objShp.Left + (objShp.Width * 0.25)
                        myTop = objShp.Top + (objShp.Height * 0.25)
                        myWidth = objShp.Width * 0.5
                        myHeight = objShp.Height * 0.5
                        myLineWeight = objShp.Line.Weight / 5
                        fontDel = 1
                    Case 2
                        myLeft = objShp.Left
                        myTop = objShp.Top
                        myWidth = objShp.Width
                        myHeight = objShp.Height
                    End Select
                Else
                    myLeft = objShp.Left
                    myTop = objShp.Top
                    myWidth = objShp.Width
                    myHeight = objShp.Height
                End If
                'If myLeft < myLeftM Or myLeftM = 0 Then myLeftM = myLeft
                'If myTop < myTopM Or myTopM = 0 Then myTopM = myTop
                objShp.Copy
                Sleep 5
                ActiveSheet.Paste
                Selection.Left = myLeft
                Selection.Top = myTop
                Selection.Width = myWidth
                Selection.Height = myHeight
                Selection.Name = Selection.Name & "a"
                If fontDel = 1 Then Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
                '�Ώۂ̃^�C�v���I�[�g�V�F�C�v(1)�̎�
                If Selection.ShapeRange.Type = 1 Then
                    '�e�L�X�g��S���������Ƀt�H���g�T�C�Y������������
                    If Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = "S" Then
                        If Selection.Width > Selection.Height Then
                            sFontSize = Selection.Height * 1.05
                        Else
                            sFontSize = Selection.Width * 1.05
                        End If
                        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = sFontSize
                    End If
                End If
                '�}���}�̐F�����̎��A�O�g�̐F�𔒂ɂ���
                If Right(Selection.Name, 5) = "Felta" And ��n���}�\�� <> 2 Then
                    Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 255)
                End If
                mygroup = mygroup & "," & Selection.Name
                
                If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                    ActiveSheet.Shapes.Range(Selection.Name).Line.Weight = myLineWeight
                    If ActiveSheet.Shapes.Range(Selection.Name).Fill.ForeColor.RGB = 1315860 Then
                        ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(255, 255, 255)
                    Else
                        ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(20, 20, 20)
                    End If
                    If ��n���}�\�� = 2 Then
                        Selection.ShapeRange.Fill.Patterned msoPatternWideDownwardDiagonal
                        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 250)
                        Selection.ShapeRange.Fill.BackColor.RGB = RGB(20, 20, 20)
                        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
                    End If
                End If
                
                'ActiveSheet.Shapes.Range(objShp.Name).Select False
                'myFlagName = ""
            End If
line20:
        Next
        g = Split(mygroup, ",")
        For i = LBound(g) To UBound(g)
            If g(i) <> "" Then
                If �[���} = Left(g(i), �[���}Len) And g(i) Like "*a*" Then
                    ActiveSheet.Shapes.Range(g(i)).Select False
                End If
            End If
        Next i
        Selection.OnAction = ""
        'Selection.Group.Select
        Selection.Left = Selection.Left + Selection.Width + Selection.Width + 6
        Selection.Name = �[���} & "_AH"
Selection.Placement = xlMove '�Z���ɍ��킹�Ĉړ��͂��邪�T�C�Y�ύX�͂��Ȃ�
'Debug.Print Round(Timer - pTime, 2)
End Function

Public Function ���`��_SH_season2(Optional �摜�� As String)
If �摜�� = "" Then �摜�� = Application.Caller
Dim pTime As Single: pTime = Timer
    Dim �[���} As String: �[���} = Left(�摜��, Len(�摜��) - 2)
    Dim �[���}Len As Long: �[���}Len = Len(�[���})
    Dim myLeft, myTop, myWidth, myHeight As Long
    Dim myLeftM As Long
    Dim myTopM As Long
    Dim myWidthM As Long
    Dim myHieghtM As Long
    Dim myLineWeight As Long
    Dim myFlag As Long
    Dim myFlagName As String
    Dim objShp As Shape
    Dim aa As Boolean
    Dim myFlagRange As Variant
    Dim c As Variant
    Dim skipFlag As Long
    Dim fontDel As Long
    On Error Resume Next
    aa = IsObject(ActiveSheet.Shapes.Range(�[���} & "_SH"))
    On Error GoTo 0
    If aa = True Then ActiveSheet.Shapes.Range(�[���} & "_SH").Delete
        
        For Each objShp In ActiveSheet.Shapes
            If �[���} = Left(objShp.Name, �[���}Len) And Not objShp.Name Like "*AH*" Then
            'Debug.Print objShp.Name
                skipFlag = 0
                If myFlagName <> "" Then
                    myFlagRange = Split(myFlagName, ",")
                    For Each c In myFlagRange
                        If c <> "" Then
                            If Left(objShp.Name, InStrRev(objShp.Name, "_")) & "e" = c Then
                                skipFlag = 1
                                Exit For
                            End If
                        End If
                    Next c
                    'kkk = InStr(myFlagName, objShp.Name)
                End If
                'If skipFlag = 1 Then GoTo line20
                
                'If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                    myFlagName = myFlagName & "," & objShp.Name & "_e"
                'End If
                
                fontDel = 0
                    myLeft = objShp.Left
                    myTop = objShp.Top
                    myWidth = objShp.Width
                    myHeight = objShp.Height
'                End If
                'If myLeft < myLeftM Or myLeftM = 0 Then myLeftM = myLeft
                'If myTop < myTopM Or myTopM = 0 Then myTopM = myTop
                objShp.Copy
                Sleep 5
                ActiveSheet.Paste
                Selection.Left = myLeft
                Selection.Top = myTop
                Selection.Width = myWidth
                Selection.Height = myHeight
                Selection.Name = Selection.Name & "s"
                mygroup = mygroup & "," & Selection.Name
                
'                If objShp.Line.ForeColor = RGB(255, 80, 80) Then
'                    ActiveSheet.Shapes.Range(Selection.Name).Line.Weight = myLineWeight
'                    If ActiveSheet.Shapes.Range(Selection.Name).Fill.ForeColor.RGB = 1315860 Then
'                        ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(255, 255, 255)
'                    Else
'                        ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(20, 20, 20)
'                    End If
'                    If ��n���}�\�� = 2 Then
'                        Selection.ShapeRange.Fill.Patterned msoPatternWideDownwardDiagonal
'                        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(20, 20, 20)
'                        Selection.ShapeRange.Fill.BackColor.RGB = RGB(255, 255, 255)
'                        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
'                    End If
'                End If
                
                'ActiveSheet.Shapes.Range(objShp.Name).Select False
                'myFlagName = ""
            End If
line20:
        Next
        g = Split(mygroup, ",")
        For i = LBound(g) To UBound(g)
            If g(i) <> "" Then
                If �[���} = Left(g(i), �[���}Len) And g(i) Like "*s" Then
                    ActiveSheet.Shapes.Range(g(i)).Select False
                End If
            End If
        Next i
        Selection.OnAction = ""
        'Selection.Group.Select
        Selection.Left = Selection.Left + Selection.Width + 3
        Selection.Name = �[���} & "_SH"
Selection.Placement = xlMove '�Z���ɍ��킹�Ĉړ��͂��邪�T�C�Y�ύX�͂��Ȃ�
Debug.Print Round(Timer - pTime, 2)
End Function


Public Function ���`��_SH()
Dim pTime As Single: pTime = Timer
    Dim �[���} As String: �[���} = Left(Application.Caller, Len(Application.Caller) - 2)
    Dim �[���}Len As Long: �[���}Len = Len(�[���})
    Dim myLeft, myTop, myWidth, myHeight As Long
    Dim myLeftM As Long
    Dim myTopM As Long
    Dim myWidthM As Long
    Dim myHieghtM As Long
    Dim myLineWeight As Long
    Dim myLineColor As Long
    Dim myFlag(5) As Long
    Dim myFlagName As String
    Dim myFlagName2 As String
    Dim objShp As Shape
    Dim aa As Boolean
    
    '������n���}����������폜
    On Error Resume Next
    aa = IsObject(ActiveSheet.Shapes.Range(�[���} & "_SH"))
    On Error GoTo 0
    If aa = True Then ActiveSheet.Shapes.Range(�[���} & "_SH").Delete
    
    
    For Each objShp In ActiveSheet.Shapes
        If �[���} = Left(objShp.Name, �[���}Len) Then
        'Debug.Print objShp.Name
            If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                myFlagName = objShp.Name
            End If
            
            For i = LBound(myFlag) To UBound(myFlag)
                myFlag(i) = 0
            Next i
            not��n�� = 0
            If objShp.Name Like "*Felt*" Then myFlag(0) = 1: not��n�� = 1
            If objShp.Name Like "*n*" Then myFlag(1) = 1: not��n�� = 1
            If objShp.Name Like "*t*" Then myFlag(2) = 1: not��n�� = 1
            If �[���} = objShp.Name Then myFlag(3) = 1: not��n�� = 1
            If myFlagName = Left(objShp.Name, Len(myFlagName)) And myFlagName <> "" Then myFlag(4) = 1: not��n�� = 1
            If not��n�� = 0 Then
                myFlagName2 = objShp.Name
                GoTo line20
            End If
            If myFlagName2 = Left(objShp.Name, Len(myFlagName2)) And myFlagName2 <> "" Then
                GoTo line20
            End If
            mygroup = mygroup & "," & Selection.Name
            GoTo line20
            
            myLeft = objShp.Left
            myTop = objShp.Top
            myWidth = objShp.Width
            myHeight = objShp.Height
            myLineWeight = objShp.Line.Weight / 6
            objShp.Duplicate.Select
            'objShp.Copy
            'ActiveSheet.Paste
            Selection.Left = myLeft
            Selection.Top = myTop
            Selection.Width = myWidth
            Selection.Height = myHeight
            Selection.Name = Selection.Name & "s"
            mygroup = mygroup & "," & Selection.Name
            If objShp.Line.ForeColor = RGB(255, 80, 80) Then
                ActiveSheet.Shapes.Range(Selection.Name).Line.Weight = myLineWeight
                If ActiveSheet.Shapes.Range(Selection.Name).Fill.ForeColor.RGB = 1315860 Then
                    ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(255, 255, 255)
                Else
                    ActiveSheet.Shapes.Range(Selection.Name).Line.ForeColor.RGB = RGB(20, 20, 20)
                End If
            End If
            '�[�����̎�Fill�̐F�ύX
            'If objShp.Name Like "*_t" Then
                'ActiveSheet.Shapes.Range(Selection.Name).Fill.ForeColor.RGB = RGB(120, 120, 200)
            'End If
            'ActiveSheet.Shapes.Range(objShp.Name).Select False
            'myFlagName = ""
        End If
line20:
    Next
    g = Split(mygroup, ",")
    For i = LBound(g) To UBound(g)
        If g(i) <> "" Then
            If �[���} = Left(g(i), �[���}Len) And g(i) Like "*s*" Then
                ActiveSheet.Shapes.Range(g(i)).Select False
            End If
        End If
    Next i
    Selection.OnAction = ""
    Selection.Group.Select
    Selection.Left = Selection.Left + Selection.Width + 3
    Selection.Name = �[���} & "_SH"
    Selection.Placement = xlMove '�Z���ɍ��킹�Ĉړ��͂��邪�T�C�Y�ύX�͂��Ȃ�
Debug.Print Round(Timer - pTime, 2)
End Function

Function �d���F�ŃZ����h��(myRow As Long, myCol As Long, �F�� As String)
        Dim clocode1 As Long        ' �F�P�i�[�p
        Dim clocode2 As Long        ' �F�Q�i�[�p
        Dim CloCode3 As Long
        
        If �F�� = "SI" Then
            'Stop
        Else
            Call �F�ϊ�(�F��, clocode1, clocode2, clofont)
            With Cells(myRow, myCol).Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 45
                .Gradient.ColorStops.Clear
                .Gradient.ColorStops.add(0).color = clocode1
                .Gradient.ColorStops.add(0.4).color = clocode1
                .Gradient.ColorStops.add(0.401).color = clocode2
                .Gradient.ColorStops.add(0.599).color = clocode2
                .Gradient.ColorStops.add(0.6).color = clocode1
                .Gradient.ColorStops.add(1).color = clocode1
            End With
            Cells(myRow, myCol).Font.color = clofont
        End If
    
End Function

Function �F�ϊ�(�F��, clocode1, clocode2, clofont)

    Set mysel = Selection
    Dim �F��a As String, �F��b As String
    Dim �ϊ��O As String
    With myBook.Sheets("color")
        Set key = .Cells.Find("ColorName", , , 1)
        �F�� = Replace(�F��, " ", "")
        If InStr(�F��, "/") = 0 Then
            �F��a = �F��
            �F��b = ""
        Else
            �F��a = Left(�F��, InStr(�F��, "/") - 1)
            �F��b = Mid(�F��, InStr(�F��, "/") + 1)
        End If
        
        If �F�� = "" Then
            clocode1 = RGB(255, 255, 255)
            clocode2 = RGB(255, 255, 255)
            clofont = RGB(255, 255, 255)
            mysel.Select
            Exit Function
        End If
        '�F�̓o�^�m�F
        �����F = �F��a
        Set ����x = .Columns(key.Column).Find(�����F, , , 1)
        If ����x Is Nothing Then GoTo errFlg
        
        �ϊ��O = ����x.Offset(0, 2)
        clocode1s = Split(�ϊ��O, ",")
        clocode1 = RGB(clocode1s(0), clocode1s(1), clocode1s(2))
        �ϊ��O = ����x.Offset(0, 3)
        clofonts = Split(�ϊ��O, ",")
        clofont = RGB(clofonts(0), clofonts(1), clofonts(2))
        
        clocode2 = clocode1
        If �F��b <> "" Then
            '�F�̓o�^�m�F
            �����F = �F��b
            Set ����x = .Columns(key.Column).Find(�����F, , , 1)
            If ����x Is Nothing Then GoTo errFlg
            
            �ϊ��O = ����x.Offset(0, 2)
            clocode2s = Split(�ϊ��O, ",")
            clocode2 = RGB(clocode2s(0), clocode2s(1), clocode2s(2))
        End If
    End With
    mysel.Select
    �F�ϊ� = clocode1
Exit Function
errFlg:
    MsgBox "�o�^����Ă��Ȃ��F " & �F��a & " ���܂�ł��܂��B�o�^���Ă��������B"
    Call �œK�����ǂ�
    With myBook.Sheets("color")
        .Select
        .Cells(.Cells(.Rows.count, key.Column).End(xlUp).Row + 1, key.Column) = �����F
    End With
    
    End
Return
End Function
Function �F�ϊ�css(�F��, clocode1, clocode2, clofont)
    Set mysel = Selection
    Dim �F��a As String, �F��b As String
    Dim �ϊ��O As String
    With myBook.Sheets("color")
        Set key = .Cells.Find("ColorName", , , 1)
        �F�� = Replace(�F��, " ", "")
        If InStr(�F��, "/") = 0 Then
            �F��a = �F��
            �F��b = ""
        Else
            �F��a = Left(�F��, InStr(�F��, "/") - 1)
            �F��b = Mid(�F��, InStr(�F��, "/") + 1)
        End If
        
        If �F�� = "" Then
            clocode1 = "FFF"
            clocode2 = "FFF"
            clofont = "000"
            mysel.Select
            Exit Function
        End If
        '�F�̓o�^�m�F
        �����F = �F��a
        Set ����x = .Columns(key.Column).Find(�����F, , , 1)
        If ����x Is Nothing Then GoTo errFlg
        
        �ϊ��O = ����x.Offset(0, 2)
        clocode1s = Split(�ϊ��O, ",")
        clocode1 = Format(Hex(clocode1s(0)), "00") & Format(Hex(clocode1s(1)), "00") & Format(Hex(clocode1s(2)), "00")
        �ϊ��O = ����x.Offset(0, 3)
        clofonts = Split(�ϊ��O, ",")
        clofont = Format(Hex(clofonts(0)), "00") & Format(Hex(clofonts(1)), "00") & Format(Hex(clofonts(2)), "00")
        
        clocode2 = clocode1
        If �F��b <> "" Then
            '�F�̓o�^�m�F
            �����F = �F��b
            Set ����x = .Columns(key.Column).Find(�����F, , , 1)
            If ����x Is Nothing Then GoTo errFlg
            
            �ϊ��O = ����x.Offset(0, 2)
            clocode2s = Split(�ϊ��O, ",")
            clocode2 = Hex(clocode2s(0)) & Hex(clocode2s(1)) & Hex(clocode2s(2))
        End If
    End With
    mysel.Select
Exit Function
errFlg:
    MsgBox "�o�^����Ă��Ȃ��F " & �F��a & " ���܂�ł��܂��B�o�^���Ă��������B"
    Call �œK�����ǂ�
    With myBook.Sheets("color")
        .Select
        .Cells(.Cells(.Rows.count, key.Column).End(xlUp).Row + 1, key.Column) = �����F
    End With
    
    End
Return
End Function
