Attribute VB_Name = "Module1"
Global Const LISTVIEW_MODE0 = "大图标"
Global Const LISTVIEW_MODE1 = "小图标"
Global Const LISTVIEW_MODE2 = "列表"
Global Const LISTVIEW_MODE3 = "详细资料"
Public fMainForm As frmMain


Sub Main()
    Dim fLogin As New frmLogin
    fLogin.Show vbModal
    If Not fLogin.OK Then
        '登录失败，退出应用程序
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

