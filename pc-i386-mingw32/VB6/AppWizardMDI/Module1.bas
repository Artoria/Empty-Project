Attribute VB_Name = "Module1"
Public fMainForm As frmMain


Sub Main()
    Dim fLogin As New frmLogin
    fLogin.Show vbModal
    If Not fLogin.OK Then
        '��¼ʧ�ܣ��˳�Ӧ�ó���
        End
    End If
    Unload fLogin


    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash


    fMainForm.Show
End Sub

