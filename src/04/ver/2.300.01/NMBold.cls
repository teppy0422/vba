VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NMBold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public �f�B���N�g�� As String
Public �u�b�N�� As String
Public �V�[�g�� As String

Public Sub Class_Initialize()
    Workbooks.Open Left(ActiveWorkbook.path, InStrRev(ActiveWorkbook.path, "\")) & nmbDirName & "\" & nmbBookName
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmbStartGyo As Long: nmbStartGyo = 2
        Dim nmbLastGyo As Long: nmbLastGyo = .Cells(.Rows.count, 1).End(xlUp).Row
        Dim nmb�^�C�g��Ramge As Range: Set nmb�^�C�g��Range = .Range(.Cells(1, 1), .Cells(1, .Cells(1, .Columns.count).End(xlToLeft).Column))
    End With
    For Each r In nmb�^�C�g��Range
        Dim nmb���i�i��Retsu As Long: If r = "���i" Then nmb���i�i��Retsu = r.Column
        Dim nmb�\��Retsu As Long: If r = "�\��" Then nmb�\��Retsu = r.Column
        Dim nmb�d���i��Retsu As Long: If r = "�i��" Then nmb�d���i��Retsu = r.Column
        Dim nmb�d���T�C�YRetsu As Long: If r = "����" Then nmb�d���T�C�YRetsu = r.Column
        Dim nmb�d���FRetsu As Long: If r = "�F" Then nmb�d���FRetsu = r.Column
        Dim nmb�d������Retsu As Long: If r = "����" Then nmb�d������Retsu = r.Column
    Next r
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmb���i�i��Range As Range: Set nmb���i�i��Range = .Range(.Cells(nmbStartGyo, nmb���i�i��Retsu), .Cells(nmbLastGyo, nmb���i�i��Retsu))
        Dim nmb�\��Range As Range: Set nmb�\��Range = .Range(.Cells(nmbStartGyo, nmb�\��Retsu), .Cells(nmbLastGyo, nmb�\��Retsu))
    End With
End Sub

Public Property Get �������o(���i�i��, �\��)
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
    For n = nmbStartGyo To nmslastgyo
        If .Cells(n, nmb���i�i��Retsu) = ���i�i�� Then
            If .Cells(n, nmb�\��Retsu) = �\�� Then
                �d������ = .Cells(n, nmb�d������Retsu).Value
                Exit Property
            End If
        End If
    Next n
    End With
End Property