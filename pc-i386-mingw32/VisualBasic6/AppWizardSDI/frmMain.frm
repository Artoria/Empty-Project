VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "工程1"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   915
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1905
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "新建"
            Object.ToolTipText     =   "新建"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "打开"
            Object.ToolTipText     =   "打开"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "保存"
            Object.ToolTipText     =   "保存"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "打印"
            Object.ToolTipText     =   "打印"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "剪切"
            Object.ToolTipText     =   "剪切"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "复制"
            Object.ToolTipText     =   "复制"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "粘贴"
            Object.ToolTipText     =   "粘贴"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "粗体"
            Object.ToolTipText     =   "粗体"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "斜体"
            Object.ToolTipText     =   "斜体"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "下划线"
            Object.ToolTipText     =   "下划线"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "左对齐"
            Object.ToolTipText     =   "左对齐"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "置中"
            Object.ToolTipText     =   "置中"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "右对齐"
            Object.ToolTipText     =   "右对齐"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2619
            Text            =   "状态"
            TextSave        =   "状态"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2013/3/7"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "1:43"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1215
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1215
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0336
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0448
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0890
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09A2
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AB4
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BC6
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CD8
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "打开(&O)..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "关闭(&C)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "全部保存(&L)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "属性(&I)"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "页面设置(&U)..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "发送(&D)..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "撤消(&U)"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "剪切(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "选择性粘贴(&S)..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "视图(&V)"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "工具栏(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "状态栏(&B)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "选项(&O)..."
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "Web 浏览器(&W)"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "选项(&O)..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "窗口(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "新建窗口(&N)"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "层叠(&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "横向平铺(&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "纵向平铺(&V)"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "排列图标(&A)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "目录(&C)"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "搜索帮助主题(&S)..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A) "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "新建"
            '应做:添加 '新建' 按钮代码。
            MsgBox "添加 '新建' 按钮代码。"
        Case "打开"
            mnuFileOpen_Click
        Case "保存"
            mnuFileSave_Click
        Case "打印"
            mnuFilePrint_Click
        Case "剪切"
            mnuEditCut_Click
        Case "复制"
            mnuEditCopy_Click
        Case "粘贴"
            mnuEditPaste_Click
        Case "粗体"
            '应做:添加 '粗体' 按钮代码。
            MsgBox "添加 '粗体' 按钮代码。"
        Case "斜体"
            '应做:添加 '斜体' 按钮代码。
            MsgBox "添加 '斜体' 按钮代码。"
        Case "下划线"
            '应做:添加 '下划线' 按钮代码。
            MsgBox "添加 '下划线' 按钮代码。"
        Case "左对齐"
            '应做:添加 '左对齐' 按钮代码。
            MsgBox "添加 '左对齐' 按钮代码。"
        Case "置中"
            '应做:添加 '置中' 按钮代码。
            MsgBox "添加 '置中' 按钮代码。"
        Case "右对齐"
            '应做:添加 '右对齐' 按钮代码。
            MsgBox "添加 '右对齐' 按钮代码。"
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    '如果这个工程没有帮助文件，显示消息给用户
    '可以在“工程属性”对话框中为应用程序设置帮助文件
    If Len(App.HelpFile) = 0 Then
        MsgBox "无法显示帮助目录，该工程没有相关联的帮助。", vbInformation, Me.Caption
    Else

    On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    '如果这个工程没有帮助文件，显示消息给用户
    '可以在“工程属性”对话框中为应用程序设置帮助文件
    If Len(App.HelpFile) = 0 Then
        MsgBox "无法显示帮助目录，该工程没有相关联的帮助。", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuWindowArrangeIcons_Click()
    '应做:添加 'mnuWindowArrangeIcons_Click' 代码。
    MsgBox "添加 'mnuWindowArrangeIcons_Click' 代码。"
End Sub

Private Sub mnuWindowTileVertical_Click()
    '应做:添加 'mnuWindowTileVertical_Click' 代码。
    MsgBox "添加 'mnuWindowTileVertical_Click' 代码。"
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    '应做:添加 'mnuWindowTileHorizontal_Click' 代码。
    MsgBox "添加 'mnuWindowTileHorizontal_Click' 代码。"
End Sub

Private Sub mnuWindowCascade_Click()
    '应做:添加 'mnuWindowCascade_Click' 代码。
    MsgBox "添加 'mnuWindowCascade_Click' 代码。"
End Sub

Private Sub mnuWindowNewWindow_Click()
    '应做:添加 'mnuWindowNewWindow_Click' 代码。
    MsgBox "添加 'mnuWindowNewWindow_Click' 代码。"
End Sub

Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewWebBrowser_Click()
    Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://www.microsoft.com"
    frmB.Show
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    '应做:添加 'mnuViewRefresh_Click' 代码。
    MsgBox "添加 'mnuViewRefresh_Click' 代码。"
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    '应做:添加 'mnuEditPasteSpecial_Click' 代码。
    MsgBox "添加 'mnuEditPasteSpecial_Click' 代码。"
End Sub

Private Sub mnuEditPaste_Click()
    '应做:添加 'mnuEditPaste_Click' 代码。
    MsgBox "添加 'mnuEditPaste_Click' 代码。"
End Sub

Private Sub mnuEditCopy_Click()
    '应做:添加 'mnuEditCopy_Click' 代码。
    MsgBox "添加 'mnuEditCopy_Click' 代码。"
End Sub

Private Sub mnuEditCut_Click()
    '应做:添加 'mnuEditCut_Click' 代码。
    MsgBox "添加 'mnuEditCut_Click' 代码。"
End Sub

Private Sub mnuEditUndo_Click()
    '应做:添加 'mnuEditUndo_Click' 代码。
    MsgBox "添加 'mnuEditUndo_Click' 代码。"
End Sub

Private Sub mnuFileExit_Click()
    '卸载窗体
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    '应做:添加 'mnuFileSend_Click' 代码。
    MsgBox "添加 'mnuFileSend_Click' 代码。"
End Sub

Private Sub mnuFilePrint_Click()
    '应做:添加 'mnuFilePrint_Click' 代码。
    MsgBox "添加 'mnuFilePrint_Click' 代码。"
End Sub

Private Sub mnuFilePrintPreview_Click()
    '应做:添加 'mnuFilePrintPreview_Click' 代码。
    MsgBox "添加 'mnuFilePrintPreview_Click' 代码。"
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "页面设置"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    '应做:添加 'mnuFileProperties_Click' 代码。
    MsgBox "添加 'mnuFileProperties_Click' 代码。"
End Sub

Private Sub mnuFileSaveAll_Click()
    '应做:添加 'mnuFileSaveAll_Click' 代码。
    MsgBox "添加 'mnuFileSaveAll_Click' 代码。"
End Sub

Private Sub mnuFileSaveAs_Click()
    '应做:添加 'mnuFileSaveAs_Click' 代码。
    MsgBox "添加 'mnuFileSaveAs_Click' 代码。"
End Sub

Private Sub mnuFileSave_Click()
    '应做:添加 'mnuFileSave_Click' 代码。
    MsgBox "添加 'mnuFileSave_Click' 代码。"
End Sub

Private Sub mnuFileClose_Click()
    '应做:添加 'mnuFileClose_Click' 代码。
    MsgBox "添加 'mnuFileClose_Click' 代码。"
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "打开"
        .CancelError = False
        'ToDo: 设置 common dialog 控件的标志和属性
        .Filter = "所有文件 (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ToDo: 添加处理打开的文件的代码

End Sub

Private Sub mnuFileNew_Click()
    '应做:添加 'mnuFileNew_Click' 代码。
    MsgBox "添加 'mnuFileNew_Click' 代码。"
End Sub

