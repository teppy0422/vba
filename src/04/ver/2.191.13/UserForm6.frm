VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "�ʒm���̎擾"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
























































































































Private Sub Label2_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton1_Click()
    Dim RAN() As String
    ReDim RAN(�Ԏ�list.ListCount - 1)
    For i = 0 To �Ԏ�list.ListCount - 1
        If �Ԏ�list.Selected(i) = True Then
            �Ԏ�str = �Ԏ�str & "_" & �Ԏ�list.List(i)
        End If
    Next i
    Unload UserForm6
    Call �A�h���X�Z�b�g(ActiveWorkbook)
    Call ie_�ʒm�����擾(�Ԏ�str)
    Call ���O�o��("test", "test", "�ʒm�����擾")
End Sub

Private Sub UserForm_Initialize()

    Set myBook = ActiveWorkbook
    
    With myBook.ActiveSheet
        Dim key As Range: Set key = .Cells.Find("key_", , , 1)
        Dim firstCol As Long: firstCol = key.Column + 1
        Dim lastCol As Long: lastCol = .UsedRange.Columns.count '.Cells(.Rows.count, �d�����ʖ�.Column).End(xlUp).Row
        Dim �Ԏ�Row As Long: �Ԏ�Row = .Cells.Find("�Ԏ�_", , , 1).Row
        Dim RAN() As String: ReDim RAN(0)
        Dim j As Long
    End With
    
    '���X�g�̐ݒ�
    With �Ԏ�list
        .RowSource = ""
        .Clear
    End With
    
    With ActiveSheet
        For X = firstCol To lastCol
            �Ԏ� = .Cells(�Ԏ�Row, X)
            flg = False
            For jj = 0 To j - 1
                If .Cells(�Ԏ�Row, X) = RAN(jj) Then
                    flg = True
                    Exit For
                End If
            Next jj
            '�V�K�ǉ�
            If flg = False Then
                ReDim Preserve RAN(j)
                RAN(j) = �Ԏ�
                j = j + 1
            End If
        Next X
    End With
    
    With �Ԏ�list
        For i = LBound(RAN) To UBound(RAN)
            If RAN(i) <> "" Then
                �Ԏ�list.AddItem RAN(i)
            End If
        Next i
    End With
    
End Sub

Private Sub �Ԏ�list_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        For i = 0 To �Ԏ�list.ListCount - 1
            �Ԏ�list.Selected(i) = False
        Next i
    End If
End Sub
