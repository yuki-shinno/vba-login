VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loginForm 
   Caption         =   "ログイン"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   OleObjectBlob   =   "loginForm.frx":0000
   StartUpPosition =   2  '画面の中央
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
    'IDとパスワードの取得
    Dim id As String
    Dim pass As String
    id = txtId.Text
    pass = txtPass.Text
    
    'IDを検索しパスワードと照合
    Dim idRow As Long
    On Error GoTo failed
    idRow = WorksheetFunction.Match(id, wsData.Columns(eColIndex.id), 0)
    If pass = wsData.Cells(idRow, eColIndex.pass) Then
        MsgBox "ログインしました", vbInformation, "成功"
        Application.Visible = True
        Unload Me
    Else
failed:
        MsgBox "ログインに失敗しました", vbCritical, "失敗"
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Dim ret As Variant
        ret = MsgBox("フォームを閉じる場合、このブックも同時に閉じられます。よろしいですか？", vbOKCancel, "確認")
        If ret = vbOK Then
            Application.DisplayAlerts = False
            ThisWorkbook.Close
            Application.DisplayAlerts = True
        ElseIf ret = vbCancel Then
            Cancel = True
        End If
    End If
End Sub
