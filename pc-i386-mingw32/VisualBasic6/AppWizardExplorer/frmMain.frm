VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "����1"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   915
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
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
         Caption         =   " �б���ͼ:"
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   4
         Tag             =   " �б���ͼ:"
         Top             =   12
         Width           =   3216
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ������ͼ:"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   " ������ͼ:"
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
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "��ǰ"
            Object.ToolTipText     =   "��ǰ"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ճ��"
            Object.ToolTipText     =   "ճ��"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ɾ��"
            Object.ToolTipText     =   "ɾ��"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "��ͼ��"
            Object.ToolTipText     =   "��ͼ��"
            ImageKey        =   "View Large Icons"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Сͼ��"
            Object.ToolTipText     =   "Сͼ��"
            ImageKey        =   "View Small Icons"
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�б�"
            Object.ToolTipText     =   "�б�"
            ImageKey        =   "View List"
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "��ϸ����"
            Object.ToolTipText     =   "��ϸ����"
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
            Text            =   "״̬"
            TextSave        =   "״̬"
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
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "��(&O)..."
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "����(&F)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "���͵�(&D)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "�½�(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "ɾ��(&D)"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "������(&M)"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "����(&I)"
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
         Caption         =   "�ر�(&C)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "����(&U)"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "����(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "ճ��(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "ѡ����ճ��(&S)..."
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "ȫѡ(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "����ѡ��(&I)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "��ͼ(&V)"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "������(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "״̬��(&B)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "��ϸ����(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "����ͼ��(&I)"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "ѡ��(&O)..."
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "Web �����(&W)"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "����(&T)"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "ѡ��(&O)..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "����(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "�½�����(&N)"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "����ƽ��(&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "����ƽ��(&V)"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "����ͼ��(&A)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Ŀ¼(&C)"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "������������(&S)..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A) "
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
    

    '���� Width ����
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    tvTreeView.Width = X
    imgSplitter.Left = X
    lvListView.Left = X + 40
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40


    '���� Top ����
  

    If tbToolBar.Visible Then
        tvTreeView.Top = tbToolBar.Height + picTitles.Height
    Else
        tvTreeView.Top = picTitles.Height
    End If

  lvListView.Top = tvTreeView.Top
    

    '���� height ����
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
        Case "����"
            'Ӧ��:��� '����' ��ť���롣
            MsgBox "��� '����' ��ť���롣"
        Case "��ǰ"
            'Ӧ��:��� '��ǰ' ��ť���롣
            MsgBox "��� '��ǰ' ��ť���롣"
        Case "����"
            mnuEditCut_Click
        Case "����"
            mnuEditCopy_Click
        Case "ճ��"
            mnuEditPaste_Click
        Case "ɾ��"
            mnuFileDelete_Click
        Case "����"
            mnuFileProperties_Click
        Case "��ͼ��"
            lvListView.View = lvwIcon
        Case "Сͼ��"
            lvListView.View = lvwSmallIcon
        Case "�б�"
            lvListView.View = lvwList
        Case "��ϸ����"
            lvListView.View = lvwReport
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    '����������û�а����ļ�����ʾ��Ϣ���û�
    '�����ڡ��������ԡ��Ի�����ΪӦ�ó������ð����ļ�
    If Len(App.HelpFile) = 0 Then
        MsgBox "�޷���ʾ����Ŀ¼���ù���û��������İ�����", vbInformation, Me.Caption
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


    '����������û�а����ļ�����ʾ��Ϣ���û�
    '�����ڡ��������ԡ��Ի�����ΪӦ�ó������ð����ļ�
    If Len(App.HelpFile) = 0 Then
        MsgBox "�޷���ʾ����Ŀ¼���ù���û��������İ�����", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuWindowArrangeIcons_Click()
    'Ӧ��:��� 'mnuWindowArrangeIcons_Click' ���롣
    MsgBox "��� 'mnuWindowArrangeIcons_Click' ���롣"
End Sub

Private Sub mnuWindowTileVertical_Click()
    'Ӧ��:��� 'mnuWindowTileVertical_Click' ���롣
    MsgBox "��� 'mnuWindowTileVertical_Click' ���롣"
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    'Ӧ��:��� 'mnuWindowTileHorizontal_Click' ���롣
    MsgBox "��� 'mnuWindowTileHorizontal_Click' ���롣"
End Sub

Private Sub mnuWindowCascade_Click()
    'Ӧ��:��� 'mnuWindowCascade_Click' ���롣
    MsgBox "��� 'mnuWindowCascade_Click' ���롣"
End Sub

Private Sub mnuWindowNewWindow_Click()
    'Ӧ��:��� 'mnuWindowNewWindow_Click' ���롣
    MsgBox "��� 'mnuWindowNewWindow_Click' ���롣"
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
    'Ӧ��:��� 'mnuViewRefresh_Click' ���롣
    MsgBox "��� 'mnuViewRefresh_Click' ���롣"
End Sub



Private Sub mnuVAIByDate_Click()
    'ToDo: ��� 'mnuVAIByDate_Click' ����
'  lvListView.SortKey = DATE_COLUMN
End Sub


Private Sub mnuVAIByName_Click()
    'ToDo: ��� 'mnuVAIByName_Click' ����
'  lvListView.SortKey = NAME_COLUMN
End Sub


Private Sub mnuVAIBySize_Click()
    'ToDo: ��� 'mnuVAIBySize_Click' ����
'  lvListView.SortKey = SIZE_COLUMN
End Sub


Private Sub mnuVAIByType_Click()
    'ToDo: ��� 'mnuVAIByType_Click' ����
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
    'Ӧ��:��� 'mnuEditInvertSelection_Click' ���롣
    MsgBox "��� 'mnuEditInvertSelection_Click' ���롣"
End Sub

Private Sub mnuEditSelectAll_Click()
    'Ӧ��:��� 'mnuEditSelectAll_Click' ���롣
    MsgBox "��� 'mnuEditSelectAll_Click' ���롣"
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'Ӧ��:��� 'mnuEditPasteSpecial_Click' ���롣
    MsgBox "��� 'mnuEditPasteSpecial_Click' ���롣"
End Sub

Private Sub mnuEditPaste_Click()
    'Ӧ��:��� 'mnuEditPaste_Click' ���롣
    MsgBox "��� 'mnuEditPaste_Click' ���롣"
End Sub

Private Sub mnuEditCopy_Click()
    'Ӧ��:��� 'mnuEditCopy_Click' ���롣
    MsgBox "��� 'mnuEditCopy_Click' ���롣"
End Sub

Private Sub mnuEditCut_Click()
    'Ӧ��:��� 'mnuEditCut_Click' ���롣
    MsgBox "��� 'mnuEditCut_Click' ���롣"
End Sub

Private Sub mnuEditUndo_Click()
    'Ӧ��:��� 'mnuEditUndo_Click' ���롣
    MsgBox "��� 'mnuEditUndo_Click' ���롣"
End Sub

Private Sub mnuFileClose_Click()
    'ж�ش���
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()
    'Ӧ��:��� 'mnuFileProperties_Click' ���롣
    MsgBox "��� 'mnuFileProperties_Click' ���롣"
End Sub

Private Sub mnuFileRename_Click()
    'Ӧ��:��� 'mnuFileRename_Click' ���롣
    MsgBox "��� 'mnuFileRename_Click' ���롣"
End Sub

Private Sub mnuFileDelete_Click()
    'Ӧ��:��� 'mnuFileDelete_Click' ���롣
    MsgBox "��� 'mnuFileDelete_Click' ���롣"
End Sub

Private Sub mnuFileNew_Click()
    'Ӧ��:��� 'mnuFileNew_Click' ���롣
    MsgBox "��� 'mnuFileNew_Click' ���롣"
End Sub

Private Sub mnuFileSendTo_Click()
    'Ӧ��:��� 'mnuFileSendTo_Click' ���롣
    MsgBox "��� 'mnuFileSendTo_Click' ���롣"
End Sub

Private Sub mnuFileFind_Click()
    'Ӧ��:��� 'mnuFileFind_Click' ���롣
    MsgBox "��� 'mnuFileFind_Click' ���롣"
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "��"
        .CancelError = False
        'ToDo: ���� common dialog �ؼ��ı�־������
        .Filter = "�����ļ� (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ToDo: ��Ӵ���򿪵��ļ��Ĵ���

End Sub

