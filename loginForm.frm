VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loginForm 
   Caption         =   "���O�C��"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   OleObjectBlob   =   "loginForm.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "loginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Enum eColIndex
    id = 1
    pass = 2
End Enum

Private Sub btnLogin_Click()
    'ID�ƃp�X���[�h�̎擾
    Dim id As String
    Dim pass As String
    id = txtId.Text
    pass = txtPass.Text
    
    'ID���������p�X���[�h�Əƍ�
    Dim idRow As Long
    On Error GoTo failed
    idRow = WorksheetFunction.Match(id, wsData.Columns(eColIndex.id), 0)
    If pass = wsData.Cells(idRow, eColIndex.pass) Then
        MsgBox "���O�C�����܂���", vbInformation, "����"
        Application.Visible = True
        Unload Me
    Else
failed:
        MsgBox "���O�C���Ɏ��s���܂���", vbCritical, "���s"
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Dim ret As Variant
        ret = MsgBox("�t�H�[�������ꍇ�A���̃u�b�N�������ɕ����܂��B��낵���ł����H", vbOKCancel, "�m�F")
        If ret = vbOK Then
            Application.DisplayAlerts = False
            ThisWorkbook.Close
            Application.DisplayAlerts = True
        ElseIf ret = vbCancel Then
            Cancel = True
        End If
    End If
End Sub
