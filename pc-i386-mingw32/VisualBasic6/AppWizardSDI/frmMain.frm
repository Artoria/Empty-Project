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
            Key             =   "�½�"
            Object.ToolTipText     =   "�½�"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "��"
            Object.ToolTipText     =   "��"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "��ӡ"
            Object.ToolTipText     =   "��ӡ"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ճ��"
            Object.ToolTipText     =   "ճ��"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "б��"
            Object.ToolTipText     =   "б��"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�»���"
            Object.ToolTipText     =   "�»���"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�����"
            Object.ToolTipText     =   "�����"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�Ҷ���"
            Object.ToolTipText     =   "�Ҷ���"
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
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "�½�(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "��(&O)..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "�ر�(&C)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "���Ϊ(&A)..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "ȫ������(&L)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "����(&I)"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "ҳ������(&U)..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "����(&D)..."
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
         Caption         =   "�˳�(&X)"
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
        Case "�½�"
            'Ӧ��:��� '�½�' ��ť���롣
            MsgBox "��� '�½�' ��ť���롣"
        Case "��"
            mnuFileOpen_Click
        Case "����"
            mnuFileSave_Click
        Case "��ӡ"
            mnuFilePrint_Click
        Case "����"
            mnuEditCut_Click
        Case "����"
            mnuEditCopy_Click
        Case "ճ��"
            mnuEditPaste_Click
        Case "����"
            'Ӧ��:��� '����' ��ť���롣
            MsgBox "��� '����' ��ť���롣"
        Case "б��"
            'Ӧ��:��� 'б��' ��ť���롣
            MsgBox "��� 'б��' ��ť���롣"
        Case "�»���"
            'Ӧ��:��� '�»���' ��ť���롣
            MsgBox "��� '�»���' ��ť���롣"
        Case "�����"
            'Ӧ��:��� '�����' ��ť���롣
            MsgBox "��� '�����' ��ť���롣"
        Case "����"
            'Ӧ��:��� '����' ��ť���롣
            MsgBox "��� '����' ��ť���롣"
        Case "�Ҷ���"
            'Ӧ��:��� '�Ҷ���' ��ť���롣
            MsgBox "��� '�Ҷ���' ��ť���롣"
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

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
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

Private Sub mnuFileExit_Click()
    'ж�ش���
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    'Ӧ��:��� 'mnuFileSend_Click' ���롣
    MsgBox "��� 'mnuFileSend_Click' ���롣"
End Sub

Private Sub mnuFilePrint_Click()
    'Ӧ��:��� 'mnuFilePrint_Click' ���롣
    MsgBox "��� 'mnuFilePrint_Click' ���롣"
End Sub

Private Sub mnuFilePrintPreview_Click()
    'Ӧ��:��� 'mnuFilePrintPreview_Click' ���롣
    MsgBox "��� 'mnuFilePrintPreview_Click' ���롣"
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "ҳ������"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'Ӧ��:��� 'mnuFileProperties_Click' ���롣
    MsgBox "��� 'mnuFileProperties_Click' ���롣"
End Sub

Private Sub mnuFileSaveAll_Click()
    'Ӧ��:��� 'mnuFileSaveAll_Click' ���롣
    MsgBox "��� 'mnuFileSaveAll_Click' ���롣"
End Sub

Private Sub mnuFileSaveAs_Click()
    'Ӧ��:��� 'mnuFileSaveAs_Click' ���롣
    MsgBox "��� 'mnuFileSaveAs_Click' ���롣"
End Sub

Private Sub mnuFileSave_Click()
    'Ӧ��:��� 'mnuFileSave_Click' ���롣
    MsgBox "��� 'mnuFileSave_Click' ���롣"
End Sub

Private Sub mnuFileClose_Click()
    'Ӧ��:��� 'mnuFileClose_Click' ���롣
    MsgBox "��� 'mnuFileClose_Click' ���롣"
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

Private Sub mnuFileNew_Click()
    'Ӧ��:��� 'mnuFileNew_Click' ���롣
    MsgBox "��� 'mnuFileNew_Click' ���롣"
End Sub

