VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim NMBdb As New Collection

Public ���i�i��nmB As String
Public �\��nmB As String
Public �i��nmB As String
Public �T�C�YnmB As String
Public �T�C�Y��nmB As String
Public �FnmB As String
Public �F��nmB As String
Public ����nmb As String
Public ����nmB As Long
Public �snmb As Long
Public ��1nmb As String
Public ��2nmb As String
Public �[�q11nmb As String
Public �[�q21nmb As String
Public �[��1nmb As String
Public �[��2nmb As String
Public Felt1nmb As String
Public Felt2nmb As String
Public titleNmb As Range

Public Function NMBset(book, sheet)
    Dim sh As Worksheet
    Set sh = Workbooks(book).Sheets(sheet)
    With Workbooks(book).Sheets(sheet)
        Dim nmbLastRow As Long
        nmbLastRow = .UsedRange.Rows.count
        Dim nmbRow As Long: nmbRow = .Cells.Find("No.").Row
        Dim nmbCol As Long: nmbCol = .Cells.Find("No.").Column
        Set titleNmb = .Range(.Cells(nmbRow, nmbCol), .Cells(nmbRow, .Columns.count))
        ���ic = titleNmb.Find("���i", , , xlWhole).Column
        �ݕ�c = titleNmb.Find("�ݕ�", , , xlWhole).Column
        �\��c = titleNmb.Find("�\��", , , xlWhole).Column
        �i��c = titleNmb.Find("�i��", , , xlWhole).Column
        �T�C�Yc = titleNmb.Find("����", , , xlWhole).Column
        �T�C�Y��c = titleNmb.Find("�T��", , , xlWhole).Column
        �Fc = titleNmb.Find("�F", , , xlWhole).Column
        ����c = titleNmb.Find("����", , , xlWhole).Column
        ����c = titleNmb.Find("����", , , xlWhole).Column
        �F��c = titleNmb.Find("�F��", , , xlWhole).Column
        ��1c = titleNmb.Find("��1", , , xlWhole).Column
        ��2c = titleNmb.Find("��2", , , xlWhole).Column
        �[�q11c = titleNmb.Find("���i11", , , xlWhole).Column
        �[�q21c = titleNmb.Find("���i21", , , xlWhole).Column
        �[��1c = titleNmb.Find("�[��1", , , xlWhole).Column
        �[��2c = titleNmb.Find("�[��2", , , xlWhole).Column
        Felt1c = titleNmb.Find("ό�1", , , xlWhole).Column
        Felt2c = titleNmb.Find("ό�2", , , xlWhole).Column
    End With

    Dim i As Long
    For i = 2 To nmbLastRow
        With New NMB
            .���i�i��nmB = sh.Cells(i, ���ic)
            .�\��nmB = sh.Cells(i, �\��c)
            .�i��nmB = sh.Cells(i, �i��c)
            .�T�C�YnmB = sh.Cells(i, �T�C�Yc)
            .�T�C�Y��nmB = sh.Cells(i, �T�C�Y��c)
            .�FnmB = sh.Cells(i, �Fc)
            .�F��nmB = sh.Cells(i, �F��c)
            .����nmb = sh.Cells(i, ����c)
            .����nmB = sh.Cells(i, ����c)
            .�snmb = sh.Cells(i).Row
            .��1nmb = sh.Cells(i, ��1c)
            .��2nmb = sh.Cells(i, ��2c)
            .�[�q11nmb = sh.Cells(i, �[�q11c)
            .�[�q21nmb = sh.Cells(i, �[�q21c)
            .�[��1nmb = sh.Cells(i, �[��1c)
            .�[��2nmb = sh.Cells(i, �[��2c)
            .Felt1nmb = sh.Cells(i, Felt1c)
            .Felt2nmb = sh.Cells(i, Felt2c)
            NMBdb.add .Self
        End With
    Next
    Set sh = Nothing
End Function

Public Property Get Self() As NMB
    Set Self = Me
End Property

Public Sub NMBseek(product, cons, ByRef found)
    For Each n In NMBdb
        If n.�\��nmB = cons Then
            If Replace(n.���i�i��nmB, " ", "") = product Then
                �i��nmB = n.�i��nmB
                �T�C�YnmB = n.�T�C�YnmB
                �T�C�Y��nmB = n.�T�C�Y��nmB
                �FnmB = n.�FnmB
                �F��nmB = n.�F��nmB
                ����nmB = n.����nmB
                ����nmb = n.����nmb
                �snmb = n.�snmb
                ��1nmb = n.��1nmb
                ��2nmb = n.��2nmb
                �[�q11nmb = n.�[�q11nmb
                �[�q21nmb = n.�[�q21nmb
                �[��1nmb = n.�[��1nmb
                �[��2nmb = n.�[��2nmb
                Felt1nmb = n.Felt1nmb
                Felt2nmb = n.Felt2nmb
                found = 1
                Exit Sub
            End If
        End If
    Next n
    '������Ȃ��ꍇ
    'Stop
    found = 0
End Sub
