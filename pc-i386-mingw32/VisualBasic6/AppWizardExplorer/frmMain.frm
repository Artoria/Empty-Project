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
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   7
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4800
      Left            =   0
      TabIndex        =   6
      Top             =   705
      Width           =   2016
      _ExtentX        =   3545
      _ExtentY        =   8467
      _Version        =   393217
      PathSeparator   =   ""
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2052
      TabIndex        =   5
      Top             =   705
      Width           =   3216
      _ExtentX        =   5662
      _ExtentY        =   8467
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4680
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   4680
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 列表视图:"
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   4
         Tag             =   " 列表视图:"
         Top             =   12
         Width           =   3216
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 树形视图:"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   " 树形视图:"
         Top             =   12
         Width           =   2016
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "返回"
            Object.ToolTipText     =   "返回"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "向前"
            Object.ToolTipText     =   "向前"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "剪切"
            Object.ToolTipText     =   "剪切"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "复制"
            Object.ToolTipText     =   "复制"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "粘贴"
            Object.ToolTipText     =   "粘贴"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "删除"
            Object.ToolTipText     =   "删除"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "属性"
            Object.ToolTipText     =   "属性"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "大图标"
            Object.ToolTipText     =   "大图标"
            ImageKey        =   "View Large Icons"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "小图标"
            Object.ToolTipText     =   "小图标"
            ImageKey        =   "View Small Icons"
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "列表"
            Object.ToolTipText     =   "列表"
            ImageKey        =   "View List"
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "详细资料"
            Object.ToolTipText     =   "详细资料"
            ImageKey        =   "View Details"
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
            TextSave        =   "1:44"
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0336
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0448
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066C
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0890
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09A2
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AB4
            Key             =   "View Details"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4788
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "打开(&O)..."
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "查找(&F)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "发送到(&D)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "重命名(&M)"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "属性(&I)"
      End
      Begin VB.Menu mnuFileBar3 
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
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "关闭(&C)"
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
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "全选(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "插入选定(&I)"
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
      Begin VB.Menu mnuListViewMode 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "排列图标(&I)"
      End
      Begin VB.Menu mnuViewBar2 
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
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim mbMoving As Boolean
Const sglSplitLimit = 500

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub


Private Sub Form_Paint()
    lvListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    Select Case lvListView.View
        Case lvwIcon
            tbToolBar.Buttons(LISTVIEW_MODE0).Value = tbrPressed
        Case lvwSmallIcon
            tbToolBar.Buttons(LISTVIEW_MODE1).Value = tbrPressed
        Case lvwList
            tbToolBar.Buttons(LISTVIEW_MODE2).Value = tbrPressed
        Case lvwReport
            tbToolBar.Buttons(LISTVIEW_MODE3).Value = tbrPressed
    End Select
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
    SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    

    '设置 Width 属性
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    tvTreeView.Width = X
    imgSplitter.Left = X
    lvListView.Left = X + 40
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40


    '设置 Top 属性
  

    If tbToolBar.Visible Then
        tvTreeView.Top = tbToolBar.Height + picTitles.Height
    Else
        tvTreeView.Top = picTitles.Height
    End If

  lvListView.Top = tvTreeView.Top
    

    '设置 height 属性
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    

    lvListView.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "返回"
            '应做:添加 '返回' 按钮代码。
            MsgBox "添加 '返回' 按钮代码。"
        Case "向前"
            '应做:添加 '向前' 按钮代码。
            MsgBox "添加 '向前' 按钮代码。"
        Case "剪切"
            mnuEditCut_Click
        Case "复制"
            mnuEditCopy_Click
        Case "粘贴"
            mnuEditPaste_Click
        Case "删除"
            mnuFileDelete_Click
        Case "属性"
            mnuFileProperties_Click
        Case "大图标"
            lvListView.View = lvwIcon
        Case "小图标"
            lvListView.View = lvwSmallIcon
        Case "列表"
            lvListView.View = lvwList
        Case "详细资料"
            lvListView.View = lvwReport
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



Private Sub mnuVAIByDate_Click()
    'ToDo: 添加 'mnuVAIByDate_Click' 代码
'  lvListView.SortKey = DATE_COLUMN
End Sub


Private Sub mnuVAIByName_Click()
    'ToDo: 添加 'mnuVAIByName_Click' 代码
'  lvListView.SortKey = NAME_COLUMN
End Sub


Private Sub mnuVAIBySize_Click()
    'ToDo: 添加 'mnuVAIBySize_Click' 代码
'  lvListView.SortKey = SIZE_COLUMN
End Sub


Private Sub mnuVAIByType_Click()
    'ToDo: 添加 'mnuVAIByType_Click' 代码
'  lvListView.SortKey = TYPE_COLUMN
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuEditInvertSelection_Click()
    '应做:添加 'mnuEditInvertSelection_Click' 代码。
    MsgBox "添加 'mnuEditInvertSelection_Click' 代码。"
End Sub

Private Sub mnuEditSelectAll_Click()
    '应做:添加 'mnuEditSelectAll_Click' 代码。
    MsgBox "添加 'mnuEditSelectAll_Click' 代码。"
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

Private Sub mnuFileClose_Click()
    '卸载窗体
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()
    '应做:添加 'mnuFileProperties_Click' 代码。
    MsgBox "添加 'mnuFileProperties_Click' 代码。"
End Sub

Private Sub mnuFileRename_Click()
    '应做:添加 'mnuFileRename_Click' 代码。
    MsgBox "添加 'mnuFileRename_Click' 代码。"
End Sub

Private Sub mnuFileDelete_Click()
    '应做:添加 'mnuFileDelete_Click' 代码。
    MsgBox "添加 'mnuFileDelete_Click' 代码。"
End Sub

Private Sub mnuFileNew_Click()
    '应做:添加 'mnuFileNew_Click' 代码。
    MsgBox "添加 'mnuFileNew_Click' 代码。"
End Sub

Private Sub mnuFileSendTo_Click()
    '应做:添加 'mnuFileSendTo_Click' 代码。
    MsgBox "添加 'mnuFileSendTo_Click' 代码。"
End Sub

Private Sub mnuFileFind_Click()
    '应做:添加 'mnuFileFind_Click' 代码。
    MsgBox "添加 'mnuFileFind_Click' 代码。"
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

