VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_91 
   Caption         =   "�ʒm���̎擾"
   ClientHeight    =   4305
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   2890
   OleObjectBlob   =   "UI_91.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_91"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



















































































































































































































Private Sub Label2_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton1_Click()
    Dim ran() As String
    ReDim ran(�Ԏ�list.ListCount - 1)
    For i = 0 To �Ԏ�list.ListCount - 1
        If �Ԏ�list.Selected(i) = True Then
            �Ԏ�str = �Ԏ�str & "_" & �Ԏ�list.List(i)
        End If
    Next i
    Unload UI_91
    Call addressSet(ActiveWorkbook)
    Call ie_�ʒm�����擾(�Ԏ�str)
    Call ���O�o��("test", "test", "�ʒm�����擾")
End Sub

Private Sub UserForm_Initialize()

    Set wb(0) = ThisWorkbook
    myIP = GetIPAddress
    addressSet wb(0)
    
    With wb(0).ActiveSheet
        Dim key As Range: Set key = .Cells.Find("���i�i��", , , 1)
        Dim firstCol As Long: firstCol = key.Column + 1
        Dim lastCol As Long: lastCol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
        Dim �Ԏ�Row As Long: �Ԏ�Row = .Cells.Find("�^��", , , 1).Row
        Dim ran() As String: ReDim ran(0)
        Dim j As Long
    End With
    
    '���X�g�̐ݒ�
    With �Ԏ�list
        .RowSource = ""
        .Clear
    End With
    
    With ActiveSheet
        For x = firstCol To lastCol
            �Ԏ� = .Cells(�Ԏ�Row, x)
            flg = False
            For jj = 0 To j - 1
                If .Cells(�Ԏ�Row, x) = ran(jj) Then
                    flg = True
                    Exit For
                End If
            Next jj
            '�V�K�ǉ�
            If flg = False Then
                ReDim Preserve ran(j)
                ran(j) = �Ԏ�
                j = j + 1
            End If
        Next x
    End With
    
    With �Ԏ�list
        For i = LBound(ran) To UBound(ran)
            �Ԏ�list.AddItem ran(i)
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
