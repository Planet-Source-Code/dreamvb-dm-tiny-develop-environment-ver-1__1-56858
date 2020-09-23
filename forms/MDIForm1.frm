VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "DM Tiny Develop Version 1"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10635
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicToolBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      ScaleHeight     =   6735
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   360
      Width           =   2835
      Begin Project1.Tray Tray1 
         Left            =   795
         Top             =   6030
         _ExtentX        =   529
         _ExtentY        =   529
         Icon            =   "MDIForm1.frx":2CFA
      End
      Begin VB.PictureBox PicShow1 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   34
         Picture         =   "MDIForm1.frx":5A04
         ScaleHeight     =   885
         ScaleWidth      =   330
         TabIndex        =   10
         ToolTipText     =   "Show Toolbox"
         Top             =   80
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox PicHideSrc 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   30
         Picture         =   "MDIForm1.frx":5D4A
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   45
         TabIndex        =   8
         Top             =   6495
         Visible         =   0   'False
         Width           =   675
      End
      Begin Project1.DevToolbar DevToolbar1 
         Height          =   2295
         Left            =   0
         TabIndex        =   7
         Top             =   1545
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4048
         Picture         =   "MDIForm1.frx":5F04
      End
      Begin Project1.DevButton DevButton 
         Height          =   300
         Index           =   0
         Left            =   15
         TabIndex        =   5
         Top             =   330
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ascii Chart"
         ForeColor       =   -2147483630
      End
      Begin VB.PictureBox PicTitleBar 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   2175
         TabIndex        =   3
         Top             =   0
         Width           =   2175
         Begin VB.PictureBox PicHideBut 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1875
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   15
            TabIndex        =   9
            ToolTipText     =   "Hide Toolbox"
            Top             =   15
            Width           =   225
         End
         Begin VB.Label lblTools 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tools"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   45
            TabIndex        =   4
            Top             =   30
            Width           =   390
         End
      End
      Begin Project1.DevButton DevButton 
         Height          =   300
         Index           =   1
         Left            =   15
         TabIndex        =   6
         Top             =   630
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Form Components"
         ForeColor       =   -2147483630
      End
      Begin Project1.DevButton DevButton 
         Height          =   300
         Index           =   2
         Left            =   15
         TabIndex        =   11
         Top             =   930
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Built in Functions"
         ForeColor       =   -2147483630
      End
      Begin Project1.DevButton DevButton 
         Height          =   300
         Index           =   3
         Left            =   15
         TabIndex        =   12
         Top             =   1230
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         ForeColor       =   -2147483630
      End
      Begin Project1.DevButton DevButton 
         Height          =   300
         Index           =   4
         Left            =   15
         TabIndex        =   13
         Top             =   3885
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         ForeColor       =   -2147483630
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   3015
      Top             =   525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3555
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6396
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":66E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":70DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7430
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7782
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8178
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":84CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":881C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8B3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7095
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18230
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_NEW"
            Object.ToolTipText     =   "New..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_OPEN"
            Object.ToolTipText     =   "Open..."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_SAVE"
            Object.ToolTipText     =   "Save.."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_CUT"
            Object.ToolTipText     =   "Cut..."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_COPY"
            Object.ToolTipText     =   "Copy..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_PASTE"
            Object.ToolTipText     =   "Paste..."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_FIND"
            Object.ToolTipText     =   "Find Text..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_RUN"
            Object.ToolTipText     =   "Run..."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_STOP"
            Object.ToolTipText     =   "Stop..."
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_FORM"
            Object.ToolTipText     =   "View Code..."
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_BACK"
            Object.ToolTipText     =   "Send To Back..."
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_FONT"
            Object.ToolTipText     =   "Bring To Front..."
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.PictureBox PicCtrlInfo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4710
         ScaleHeight     =   285
         ScaleWidth      =   2910
         TabIndex        =   14
         Top             =   15
         Width           =   2910
         Begin VB.Label lblTextPos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2040
            TabIndex        =   17
            Top             =   45
            Width           =   225
         End
         Begin VB.Label lblSize 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1650
            TabIndex        =   16
            Top             =   45
            Width           =   225
         End
         Begin VB.Label lblPosition 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   15
            Top             =   45
            Width           =   225
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   1290
            Picture         =   "MDIForm1.frx":8E30
            Top             =   15
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   75
            Picture         =   "MDIForm1.frx":8F7A
            Top             =   15
            Width           =   240
         End
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuopen 
         Caption         =   "&Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnubalnk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save Project"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnublank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnumake 
         Caption         =   "&Make Exe"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnublank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuall 
         Caption         =   "Select &All"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mnublank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnufront 
         Caption         =   "&Bring to Front"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBack 
         Caption         =   "&Send to Back"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnublank7 
         Caption         =   "-"
      End
      Begin VB.Menu mnufind 
         Caption         =   "&Find Text..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnureplace 
         Caption         =   "&Replace Text"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnugoto 
         Caption         =   "&Goto..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnubalnk8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconvert 
         Caption         =   "Con&vert"
         Begin VB.Menu mnuupper 
            Caption         =   "Uppercase"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnulower 
            Caption         =   "&Lowercase"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuinvert 
            Caption         =   "&Invert Case"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuspacetab 
            Caption         =   "&Space to Tab"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnutpsapace 
            Caption         =   "&Tabs tp Space"
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu mnuinsert 
         Caption         =   "&Insert"
         Begin VB.Menu mnuDate 
            Caption         =   "System &Date"
         End
         Begin VB.Menu mnuTime 
            Caption         =   "System &Time"
         End
         Begin VB.Menu mnucompname 
            Caption         =   "Computer Name"
         End
         Begin VB.Menu mnuusername 
            Caption         =   "User Name"
         End
         Begin VB.Menu mnublank9 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSysEnv 
            Caption         =   "Environment Variable"
            Begin VB.Menu MnuEnvVars 
               Caption         =   ""
               Index           =   0
            End
         End
         Begin VB.Menu mnucodeHelpers 
            Caption         =   "Code Template Helpers"
            Begin VB.Menu mnucodeHelpersA 
               Caption         =   " "
               Index           =   0
            End
         End
      End
   End
   Begin VB.Menu mnuView1 
      Caption         =   "&View"
      Begin VB.Menu mnugrid 
         Caption         =   "&Grid"
         Checked         =   -1  'True
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View Code"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnutoolbox 
         Caption         =   "Tool Bo&x..."
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu MnuEnv 
         Caption         =   "&Environment Options"
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plug-ins"
      Begin VB.Menu mnuPlgName 
         Caption         =   "p"
         Index           =   0
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodeView As Boolean, m_ShowControlBox As Boolean, DlgDesignTime As Boolean
Public nGrid As Boolean
Dim OldWidth As Long
Dim OldWinState As Integer



Enum ToolsOption
    ACharSet
    [WinFormTools]
    [DevFunctions]
    [LanFunctions]
    [CodeTemplates]
End Enum

Private DevBarOption As ToolsOption

Private Sub ShowInfoLable(mShow As Boolean)
    lblTextPos.Visible = Not mShow
    lblTextPos.Left = Image1(0).Left
    Image1(0).Visible = mShow
    Image1(1).Visible = mShow
    lblSize.Visible = mShow
    lblPosition.Visible = mShow
End Sub

Public Sub GetLastScriptError()
Dim sErrorText As String
    If dScript.ErrScript.Number = 0 Then Exit Sub
    
    sErrorText = dScript.ErrScript.Source & " '" & dScript.ErrScript.Number & "':" & vbCrLf
    sErrorText = sErrorText & vbCrLf
    sErrorText = sErrorText & dScript.ErrScript.Description & vbCrLf _
    & "Line: " & dScript.ErrScript.Line & "    Column :" & dScript.ErrScript.Column
    mGoto = dScript.ErrScript.Line
    frmDebug.lblError.Caption = sErrorText
    sErrorText = ""
    IdeStop
    frmDebug.DrawLine
    frmDebug.Show vbModal
    If ButtonPressed = 1 Then
        mnuView_Click
        HighLightLine CLng(mGoto - 1), frmCode.txtCode
    End If
    mGoto = 0
    
End Sub

Private Sub ShowToolBox(mShowBox As Boolean)
    I = 0
    If mShowBox Then
        PicShow1.Visible = False
        For I = 0 To DevButton.Count - 1
            DevButton(I).Visible = True
        Next
        PicTitleBar.Visible = True: DevToolbar1.Visible = True
        PicToolBar.Width = OldWidth
        m_ShowControlBox = False
    Else
        OldWidth = PicToolBar.Width
        PicShow1.Visible = True
        'PicShow1.Top = 80: PicShow1.Left = 34
        PicTitleBar.Visible = False: DevToolbar1.Visible = False
        For I = 0 To DevButton.Count - 1
            DevButton(I).Visible = False
        Next
        PicToolBar.Width = 400
        m_ShowControlBox = True
    End If
End Sub

Private Sub DoHideButtonState(mState As Integer)
    PicHideBut.Cls
    TransparentBlt PicHideBut.hdc, 0, 0, 15, 13, PicHideSrc.hdc, mState * 15, 0, 15, 13, RGB(255, 0, 255)
    PicHideBut.Refresh
End Sub
Private Sub DrawEffect(Direction As Integer)
    If Direction = 1 Then
        UserControl.Line (UserControl.ScaleWidth, 0)-(0, 0), vbWhite
        UserControl.Line (0, UserControl.ScaleHeight)-(0, -1), vbWhite
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), vbApplicationWorkspace
        UserControl.Line (UserControl.ScaleWidth, UserControl.ScaleHeight - 1)-(-1, UserControl.ScaleHeight - 1), vbApplicationWorkspace
    Else
        UserControl.Line (UserControl.ScaleWidth, 0)-(0, 0), vbApplicationWorkspace
        UserControl.Line (0, UserControl.ScaleHeight)-(0, -1), vbApplicationWorkspace
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), vbWhite
        UserControl.Line (UserControl.ScaleWidth, UserControl.ScaleHeight - 1)-(-1, UserControl.ScaleHeight - 1), vbWhite
    End If
End Sub

Private Sub LoadinPlugins()
Dim PlgInfo As String, PlgCount As Integer, PlgStr As String, Counter As Integer
Dim PlgArg As Variant, PlgPath As String
Dim FoundPlugin As Integer

    FoundPlugin = 0
    DevXML.XMLReset
    DevXML.XMLLoadFormFile ApplicationPath & "\Plugins.xml"
    PlgCount = CInt(DevXML.GetSelectionValue(DevXML.XMLData, "Count", "value", "0"))
    If PlgCount = 0 Then DevXML.XMLReset: Exit Sub
    
    PlgInfo = DevXML.GetSelection("Plugins")
    If DevXML.HasError Then DevXML.XMLReset: Exit Sub
    
    For Counter = 1 To PlgCount
        PlgArg = Split(DevXML.GetSelectionValue(PlgInfo, "Plugin" & CStr(Counter), "Info"), ",")
        If UBound(PlgArg) <> 2 Then
            Exit For
        ElseIf Len(PlgArg(0)) = 0 Then Exit For
        ElseIf Len(PlgArg(1)) = 0 Then Exit For
        ElseIf Len(PlgArg(2)) = 0 Then Exit For
        ElseIf InStr(1, PlgArg(2), "{PLUG_PATH}", vbTextCompare) Then
            PlgPath = Replace(PlgArg(2), "{PLUG_PATH}", ApplicationPath & "plug-ins")
        Else
            PlgPath = PlgArg(2)
        End If
        
        If Not RegisterActiveX(GetShortPath(PlgPath), Register) Then
            MsgBox "There was an error while Registering: " _
            & vbCrLf & vbCrLf & "Plugin : " & PlgArg(2) _
            & vbCrLf & vbCrLf & "Plugin Interface Name : " & PlgArg(1), vbExclamation, Err.Description
        Else
            FoundPlugin = FoundPlugin + 1
            Load mnuPlgName(FoundPlugin)
            mnuPlgName(FoundPlugin - 1).Caption = PlgArg(0)
            
            ReDim Preserve TPlugin.PlugInterFace(FoundPlugin)
            ReDim Preserve TPlugin.PlugFileName(FoundPlugin)
            TPlugin.PlugInterFace(FoundPlugin) = PlgArg(1)
            TPlugin.PlugFileName(FoundPlugin) = PlgPath
        End If
    Next
    
    Unload mnuPlgName(FoundPlugin)
    
    Erase PlgArg
    PlgInfo = ""
    PlgPath = ""
    PlgStr = ""
    PlgCount = 0
    Counter = 0
    DevXML.XMLReset
End Sub
Private Sub SetupToolbarItems()
Dim vList As Variant, sLine As String, bLine As String, I As Integer, J As Integer

    If (DevBarOption = -1) Then Exit Sub
    DevToolbar1.ResetButton
    DevToolbar1.SetupControlBar
    
    Select Case DevBarOption
        Case ACharSet
            For J = 32 To 126
                If J <= 99 Then
                    DevToolbar1.AddButton "Char     " & CStr(J) & "           '" & Chr(J) & "'", 5, Chr(J)
                Else
                    DevToolbar1.AddButton "Char    " & CStr(J) & "          '" & Chr(J) & "'", 5, Chr(J)
                End If
            Next
            DoEvents
            J = 0
            
        Case [WinFormTools]
            DevToolbar1.AddButton "Picture Box", 0, "T_IMAGE"
            DevToolbar1.AddButton "Command Button", 2, "T_BUTTON"
            DevToolbar1.AddButton "Label", 3, "T_LABEL"
            DevToolbar1.AddButton "Text Box", 4, "T_TEXT"
            DevToolbar1.AddButton "List Box", 1, "T_LIST"
            DevToolbar1.DrawToolBar
        
        Case [DevFunctions] ' in Built functions
            StrB = OpenFile(Function_List & "vbScript\Functions2.ref")
            vList = Split(StrB, vbCrLf)
            For I = LBound(vList) To UBound(vList)
                If Len(vList(I)) > 0 Then
                    sLine = vList(I)
                    DevToolbar1.AddButton Space(4) & sLine, 6, sLine
                End If
            Next
            DoEvents
            StrB = ""
            I = 0
            Erase vList
        Case [LanFunctions]
            If LCase(TProject.ProgLan) = "vbscript" Then
                StrB = OpenFile(Function_List & "vbScript\Functions1.ref")
            Else
                StrB = OpenFile(Function_List & "JScript\jFunctions1.ref")
            End If
            vList = Split(StrB, vbCrLf)
            For I = LBound(vList) To UBound(vList)
                If Len(vList(I)) > 0 Then
                    sLine = vList(I)
                    DevToolbar1.AddButton Space(4) & sLine, 6, sLine
                End If
            Next
            DoEvents
            StrB = ""
            I = 0
            Erase vList
        Case [CodeTemplates]
            For I = 0 To CodeHelperTemplate.CodehelperCount
                DevToolbar1.AddButton CodeHelperTemplate.CodehelperName(I), 7, CodeHelperTemplate.CodehelperFilePath(I)
            Next
            I = 0
    End Select
    DevToolbar1.DrawToolBar
    
End Sub

Private Sub FillCodeHelperMenu()
Dim I As Integer
    If (CodeHelperTemplate.CodehelperCount <= 0) Then
        mnucodeHelpersA(0).Caption = "-"
        Exit Sub
    End If
    ' Unload All the menus first
    If mnucodeHelpersA.Count > 1 Then
        For I = 1 To mnucodeHelpersA.Count - 1
            Unload mnucodeHelpersA(I)
        Next
    End If
    ' Add the code helper template menu to the menu
    For I = 0 To CodeHelperTemplate.CodehelperCount
        Load mnucodeHelpersA(I + 1)
        mnucodeHelpersA(I).Caption = CodeHelperTemplate.CodehelperName(I)
    Next
    Unload mnucodeHelpersA(I)
    I = 0
End Sub

Private Sub FillEnvMenu()
Dim I As Integer, iPos As Integer, StrEnvName As String
    Do
        I = I + 1
        StrEnvName = Environ(I)
        iPos = InStr(1, StrEnvName, "=", vbBinaryCompare)
        If iPos > 0 Then
            Load MnuEnvVars(I)
            MnuEnvVars(I - 1).Caption = Mid(StrEnvName, 1, iPos - 1)
        End If
        DoEvents
    Loop Until LenB(StrEnvName) = 0
    Unload MnuEnvVars(I - 1)
    I = 0: iPos = 0: StrEnvName = ""
End Sub

Private Sub CleanUpEnd()
    Unload FrmAbout
    Unload frmCode
    Unload frmmenu
    Unload frmOptions
    Unload frmproject
    Unload frmWorkArea
    Unload MDIForm1
End Sub

Private Sub EnableDisableMenu()
    Toolbar1.Buttons(5).Enabled = frmCode.EnableCutCopy
    Toolbar1.Buttons(6).Enabled = frmCode.EnableCutCopy
    mnucut.Enabled = frmCode.EnableCutCopy
    mnucopy.Enabled = frmCode.EnableCutCopy
    mnudelete.Enabled = frmCode.EnableCutCopy
    mnuupper.Enabled = frmCode.EnableCutCopy
    mnulower.Enabled = frmCode.EnableCutCopy
    mnuinvert.Enabled = frmCode.EnableCutCopy
    mnuspacetab.Enabled = frmCode.EnableCutCopy
    mnutpsapace.Enabled = frmCode.EnableCutCopy
End Sub

Sub SetupScript()
    Set dScript.DialogObject = Nothing
    dScript.DialogStrName = ""
    dScript.mLanguage = ""
    dScript.mLanguage = TProject.ProgLan
    dScript.DialogStrName = "Dialog"
    Set dScript.DialogObject = frmWorkArea
    dScript.SetupControl
End Sub
Public Sub IdeStop()
    Toolbar1.Buttons(13).Enabled = True ' enable code view button
    DlgDesignTime = True ' we are in design time
    inIde = True ' in design mode
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
    Toolbar1.Buttons(11).Enabled = False
    mnunew.Enabled = True
    mnuopen.Enabled = True
    mnusave.Enabled = True
    mnumake.Enabled = True
    mnufront.Enabled = False
    mnuBack.Enabled = False
    mnuView.Enabled = True
    mnugrid.Enabled = True
    mnutoolbox.Enabled = True
    PicToolBar.Visible = True ' show the toolbox
    ' restore the forms data
    RestoreData frmWorkArea
    frmWorkArea.Cls
    Set frmWorkArea.Picture = Nothing
    DrawGrid frmWorkArea, , mnugrid.Checked
    Set dScript.DialogObject = Nothing
    dScript.Reset
End Sub

Private Sub SetupIDE()
    If CodeView Then mnuView_Click

    If LCase(TProject.ProgLan) = "vbscript" Then
        PhaseInsightList "vb.dm" ' load the insight list for the editor
        LanComment = "'" ' vbscript comment
        DevButton(3).Caption = "VB Script Functions"
        DevButton(4).Caption = "VB Script Code Helpers"
    Else
        PhaseInsightList "java.dm" ' load the insight list for the editor
        LanComment = "//" ' javascript comment
        DevButton(3).Caption = "Java Script Functions"
        DevButton(4).Caption = "Java Script Code Helpers"
    End If
    
    PhaseHelperList TProject.ProgLan ' Load in code template Helper list
    FillCodeHelperMenu ' Load in the code helper templates menu
    
    DevButton(4).Visible = True
    
    FirstTimeLoad = False
    ShowToolBox True
    LoadForm frmWorkArea, ProjectFolder & TProject.FormFile
    
    MDIForm1.Caption = "DM Tiny Develop Version 1 - " & TProject.ProjectTitle
    frmWorkArea.Visible = True
    frmWorkArea.devToolTip1.Visible = False
    
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
    Toolbar1.Buttons(13).Enabled = True
    
    mnuedit.Enabled = True
    mnuView1.Enabled = True
    frmCode.Visible = False
    mnusave.Enabled = True
    mnumake.Enabled = True
    mnuPlugins.Enabled = True
    
    DlgDesignTime = True
    frmCode.txtCode = OpenFile(ProjectFolder & TProject.UnitFile)

    ' setup the script control
    SetupScript
    
    DevButton_DevButtonMouseUp 4, 1, 0, 0, 0
    DevButton_DevButtonMouseUp 1, 0, 0, 0, 0
    
    PicCtrlInfo.Visible = True
    MDIForm1.lblPosition.Caption = "0, 0"
    MDIForm1.lblSize.Caption = "0, 0"
End Sub
Private Sub UnloadAllControls()
On Error Resume Next
Dim I As Integer
    For I = 1 To frmWorkArea.CmdBut.Count - 1
        Unload frmWorkArea.CmdBut(I)
    Next
    I = 0
    For I = 1 To frmWorkArea.PicImg.Count - 1
        Unload frmWorkArea.PicImg(I)
    Next
    I = 0
    For I = 1 To frmWorkArea.lblA.Count - 1
        Unload frmWorkArea.lblA(I)
    Next
    I = 0
    For I = 1 To frmWorkArea.txtA.Count - 1
        Unload frmWorkArea.txtA(I)
    Next
    I = 0
    For I = 1 To frmWorkArea.lstA.Count - 1
        Unload frmWorkArea.lstA(I)
    Next
    I = 0
    
    frmWorkArea.HideSelection
End Sub



Private Sub DevButton_DevButtonMouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MDIForm_MouseMove 1, 0, 0, 0
End Sub

Private Sub DevButton_DevButtonMouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DevBarOption = Index
    MDIForm_Resize
End Sub

Private Sub DevToolbar1_DevToolBarMouseUp(Button As Integer, Index As Integer, Key As String)
    Select Case DevBarOption
        Case ACharSet
            If Not CodeView Then Exit Sub
           frmCode.txtCode.SelText = DevToolbar1.ButtonKey(Index)
           frmCode.SetFocus
        Case [WinFormTools]
            If Not DlgDesignTime Then Exit Sub
            tAddControl frmWorkArea, DevToolbar1.ButtonKey(Index)
        Case [DevFunctions]
            If Not CodeView Then Exit Sub
            frmCode.txtCode.SelText = DevToolbar1.ButtonKey(Index)
            frmCode.txtCode.SetFocus
        Case [LanFunctions]
            If Not CodeView Then Exit Sub
            frmCode.txtCode.SelText = DevToolbar1.ButtonKey(Index)
            frmCode.txtCode.SetFocus
        Case [CodeTemplates]
            If Not CodeView Then Exit Sub
            frmCode.txtCode.SelText = OpenFile(DevToolbar1.ButtonKey(Index))
            frmCode.txtCode.SetFocus
    End Select
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
    ' main loading part of the program
    
    MDIForm1.MousePointer = vbHourglass
    ApplicationPath = FixPath(App.Path)
    TemplatePath = GetShortPath(ApplicationPath & "Project Templates\")

    DataPath = ApplicationPath & "Data\"
    Function_List = DataPath & "Reference\"

    If LenB(Dir(DataPath)) = 0 Then
        ' If no data path is found above then we must exit
        End
    End If
    
    ProjectFolder = ApplicationPath & "Projects\" ' projects path
    inSightHelperList = ApplicationPath & "data\CodeInsight\" ' code insight path
    App_ConfigFile = ApplicationPath & "config.xml"
    
    If Not IsFileHere(App_ConfigFile) Then
        ' if no config file is found we just create the default one
        SaveDefaultSettings
        WriteXMLIni
    End If

    frmOptions.LoadConfig ' Load the programs config file
    FillEnvMenu ' Load in Environment Variable menu
    LoadinPlugins ' Load in the Plug-ins menu items
    If Not FindDir(ProjectFolder) Then MkDir ProjectFolder
    
    Modified = False
    CodeView = False
    DlgDesignTime = False
    FirstTimeLoad = True
    lblTextPos.Visible = False
    
    ReDim Preserve ProjectData.mCommandButton(0)
    ReDim Preserve ProjectData.mlabel(0)
    ReDim Preserve ProjectData.mPictureBox(0)
    ReDim Preserve ProjectData.mTextBox(0)
    ReDim Preserve ProjectData.nListBox(0)
    
    ' to be set as default in the menu latter
    mnuedit.Enabled = False
    mnuView1.Enabled = False
    mnufind.Enabled = False
    mnugoto.Enabled = False
    mnureplace.Enabled = False
    mnuconvert.Enabled = False
    mnuinsert.Enabled = False
    mnuPlugins.Enabled = False
    
    PicCtrlInfo.Visible = False
    
    PicTitleBar.Width = (PicToolBar.ScaleWidth)
    PicHideBut.Left = (PicTitleBar.ScaleWidth - PicHideBut.Width - 100)
    DevBarOption = -1
    FlatBorder PicTitleBar.hwnd
    
    DevButton_DevButtonMouseUp 1, 1, 0, 0, 0
    MDIForm1.MousePointer = vbDefault
    
    nGrid = AppConfig.ShowGrid
    mnugrid_Click
    DoHideButtonState 2
    ShowToolBox False
    
    DoEvents
End Sub

Public Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DevToolbar1.Visible Then DevToolbar1.HideFocus
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    DevToolbar1.Height = MDIForm1.ScaleHeight - DevToolbar1.Top - DevButton(4).Height
    DevButton(4).Top = (MDIForm1.ScaleHeight - DevButton(4).Height + 40)
    SetupToolbarItems
    ' Hi all this a little bug someone in this code. if you minsize the program to the tray then maxsize agian
    'DevButton(4).Top will be stock in the center also DevToolbar1 height is not correct. just hit one of the buttons this will resize it back then
    ' I tryed all ways and can't fix the damm thing doing my head in.
    ' anyway if anyone can help please let me know.
    
    ' code below to be replaces when I find a fix for the above
    If Me.WindowState = 1 Then
        Me.WindowState = 0
        Tray1.Visible = True
        Tray1.ToolTip = "Restore"
        MDIForm1.Visible = False
    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Tray1.Visible = False
    Tray1.ToolTip = ""
    CleanUpEnd
End Sub

Private Sub mnuabout_Click()
    FrmAbout.Show vbModal, MDIForm1
End Sub

Private Sub mnuall_Click()
    EditMenu nSelectAll, frmCode.txtCode
    EnableDisableMenu
End Sub

Private Sub mnuback_Click()
    TheObjectName.ZOrder vbSendToBack
End Sub

Private Sub mnucodeHelpersA_Click(Index As Integer)
    frmCode.txtCode.SelText = OpenFile(CodeHelperTemplate.CodehelperFilePath(Index))
End Sub

Private Sub mnucompname_Click()
    frmCode.txtCode.SelText = LanComment & SysComputerName
End Sub

Private Sub mnucopy_Click()
    If ObjectSelected = iCodeView Then ' Code view mode
        EditMenu nCopy, frmCode.txtCode
    End If
End Sub

Private Sub mnucut_Click()
    If ObjectSelected = iCodeView Then ' Code view mode
        EditMenu nCut, frmCode.txtCode
        EnableDisableMenu
    Else
        frmWorkArea.HideSelection ' hide objects selection
        Unload TheObjectName ' unload the object
    End If
End Sub

Private Sub mnuDate_Click()
    frmCode.txtCode.SelText = LanComment & Date
End Sub

Private Sub mnudelete_Click()
On Error Resume Next
    If ObjectSelected = iCodeView Then ' Code view mode
        EditMenu nDelete, frmCode.txtCode
        EnableDisableMenu
    Else
        frmWorkArea.HideSelection ' hide objects selection
        Unload TheObjectName ' unload the object
    End If
End Sub

Private Sub mnuenv_Click()
    frmOptions.Show vbModal, MDIForm1
End Sub

Private Sub MnuEnvVars_Click(Index As Integer)
MsgBox TProject.ProgLan

    If LCase(TProject.ProgLan) = "vbscript" Then
        frmCode.txtCode.SelText = Chr(34) & MnuEnvVars(Index).Caption & Chr(34)
    Else
        frmCode.txtCode.SelText = "'" & MnuEnvVars(Index).Caption & "'"
    End If
End Sub

Private Sub mnuexit_Click()
    MDIForm_Unload 0
End Sub

Private Sub mnufind_Click()
    frmFind.Show , MDIForm1
    frmCode.txtCode.SetFocus
End Sub

Private Sub mnufront_Click()
    TheObjectName.ZOrder vbBringToFront
End Sub

Private Sub mnugoto_Click()
    frmGoto.Show vbModal, MDIForm1
    If ButtonPressed = 0 Then Exit Sub ' cancel button was pressed
    
    Select Case TSelectionType
        Case 0 ' goto start of code top line
            frmCode.txtCode.SelStart = 0
            frmCode.txtCode.SetFocus
        Case 1 ' goto bottom of code last line
            frmCode.txtCode.SelStart = Len(frmCode.txtCode.Text)
            frmCode.txtCode.SetFocus
        Case 2 ' goto a selection in the code
            frmCode.txtCode.SelStart = frmCode.txtCode.SelStart + mGoto
        Case 3
            GotoLine CLng(mGoto - 1), frmCode.txtCode
    End Select
End Sub

Public Sub mnugrid_Click()
    DrawGrid frmWorkArea, AppConfig.GridColor, nGrid
    mnugrid.Checked = nGrid
    nGrid = Not nGrid
End Sub

Private Sub mnuinvert_Click()
    frmCode.txtCode.SelText = Invert(frmCode.txtCode.SelText)
End Sub

Private Sub mnulower_Click()
    frmCode.txtCode.SelText = LCase(frmCode.txtCode.SelText)
End Sub

Private Sub mnumake_Click()
Dim nDebugPath As String, sBuffer1 As String, sBuffer2 As String, iFile As Long, NewExe As String, _
TheHeadInfo As String, ExeHeadFile As String, sHead As String, manifest As String


    nDebugPath = ProjectFolder & "debug"
    NewExe = FixPath(nDebugPath) & TProject.ProjectTitle & ".exe"
    ExeHeadFile = ApplicationPath & "exehead\exehead.exe"
    
    If Not IsFileHere(ExeHeadFile) Then
        MsgBox "Compile Error unable to link code." _
        & vbCrLf & vbCrLf & ExeHeadFile & " was not found", vbCritical, "File not Found"
        Exit Sub
    End If

    If IsFileHere(NewExe) Then Kill NewExe
    If Not FindDir(nDebugPath) Then MkDir nDebugPath
    
    mnusave_Click
    
    sBuffer1 = Encrypt(RemoveComments(OpenFile(ProjectFolder & TProject.UnitFile)))
    sBuffer2 = OpenFile(ProjectFolder & TProject.FormFile)
    
    MakeExe.Win32CodeData = sBuffer1
    MakeExe.Win32FormData = sBuffer2
    MakeExe.Win32Lan = TProject.ProgLan
    
    sBuffer1 = ""
    sBuffer2 = ""
    
    manifest = OpenFile(TemplatePath & "manifest.xml")
    manifest = Replace(manifest, "{description}", TProject.ProjectTitle)
    
    iFile = FreeFile
    
    FileCopy ExeHeadFile, NewExe
    sHead = "<DATA>" & Chr(5)
    Open NewExe For Binary As #iFile
        Put #iFile, LOF(iFile), sHead
        Put #iFile, LOF(iFile) + 1, MakeExe
    Close #iFile
    
    WriteToFile FixPath(nDebugPath) & TProject.ProjectTitle & ".exe.manifest", manifest
    sHead = ""
    manifest = ""
    ExeHeadFile = ""
    
    MsgBox "Your Appliaction has now been compiled to :" _
    & vbCrLf & NewExe, vbInformation
    NewExe = ""
    
End Sub

Private Sub mnunew_Click()
    frmproject.Show vbModal, MDIForm1
    If ButtonPressed = 0 Then Exit Sub
    
    If FindDir(ProjectFolder) Then
        MsgBox "The project name you named already exsits." _
        & vbCrLf & vbCrLf & "Please choose a different name", vbInformation, Me.Caption
        Exit Sub
    Else
        MkDir ProjectFolder
        CreateProject
        If Not OpenProject(ProjectFolder & ProjectName & ".proj") Then
            MsgBox "The project can't be opened", vbCritical, "Unable To Load Project"
            Exit Sub
        Else
            UnloadAllControls
            SetupIDE
        End If
    End If
    
End Sub

Private Sub mnuopen_Click()
On Error GoTo CanError

    With CDialog
        .CancelError = True
        .DialogTitle = "Open Project"
        .Filter = "Project Files(*.proj)|*.proj|"
        .Filename = ""
        .InitDir = ProjectFolder
        .ShowOpen
        If Len(.Filename) = 0 Then Exit Sub
        
        If Not GetFileExt(.Filename) = "proj" Then
            MsgBox "There was an error while trying to open file.", vbInformation, "Unable to Open File"
            Exit Sub
        ElseIf Not OpenProject(.Filename) Then
            MsgBox "There was an error while trying to open file.", vbInformation, "Unable to Open File"
            Exit Sub
        Else
            ProjectFolder = GetAbsPath(.Filename)
            UnloadAllControls
            SetupIDE
        End If
    End With
    
CanError:
    If Err = cdlCancel Then Exit Sub
    
End Sub

Private Sub mnupaste_Click()
    If ObjectSelected = iCodeView Then ' Code view mode
        EditMenu nPaste, frmCode.txtCode
    End If
End Sub

Private Sub mnuPlgName_Click(Index As Integer)
On Error GoTo PlugErr
Dim clsPlg As New clsPlugIFace
Dim PlgObject As Object

    Set PlgObject = CreateObject(TPlugin.PlugInterFace(Index + 1))
    Set clsPlg.DevCodeWindow = frmCode
    Set clsPlg.DevIDE = MDIForm1
    
    PlgObject.RunPlugin clsPlg
    Set PlgObject = Nothing
    Set clsPlg.DevCodeWindow = Nothing
    Set clsPlg.DevIDE = Nothing
    Exit Sub
PlugErr:
    If Err Then
        MsgBox Err.Description
    End If
    
End Sub

Private Sub mnusave_Click()
    On Error Resume Next
    ' A little bug fix not perfect like but seems to work
    ' had to add the loop below because if you deleted
    ' a Form object eg button and save the old button will still be there.
    ' not sure why it does it seems VB does not like saveing or deleteing files
    ' If you can find a different way please add it in or let me know.
    Dim L As Integer
    For L = 0 To 1
        DeleteFile ProjectFolder & TProject.FormFile
        DeleteFile ProjectFolder & TProject.UnitFile
        TUnitSrc = frmCode.txtCode.Text
        SaveProject frmWorkArea
    Next
    
    L = 0
End Sub

Private Sub mnuspacetab_Click()
    frmCode.txtCode.SelText = ReplaceChr(frmCode.txtCode.SelText, Chr(32), Chr(9))
End Sub

Private Sub mnuTime_Click()
    frmCode.txtCode.SelText = LanComment & Time
End Sub

Private Sub mnutoolbox_Click()
    If FirstTimeLoad Then Exit Sub
    ShowToolBox m_ShowControlBox
End Sub

Private Sub mnutpsapace_Click()
    frmCode.txtCode.SelText = ReplaceChr(frmCode.txtCode.SelText, Chr(9), Chr(32))
End Sub

Private Sub mnuupper_Click()
    frmCode.txtCode.SelText = UCase(frmCode.txtCode.SelText)
End Sub

Private Sub mnuusername_Click()
    frmCode.txtCode.SelText = LanComment & GetUser
End Sub

Private Sub mnuView_Click()
    If mnuView.Caption = "&View Form Designer" Then
        mnuView.Caption = "&View Code"
    Else
        mnuView.Caption = "&View Form Designer"
    End If
    
    CodeView = Not CodeView
    DlgDesignTime = Not CodeView
    frmCode.Visible = CodeView
    frmWorkArea.Visible = Not CodeView
    
    Toolbar1.Buttons(5).Enabled = False ' disable cut
    Toolbar1.Buttons(6).Enabled = False ' disable copy
    Toolbar1.Buttons(8).Enabled = CodeView
    
    ShowInfoLable Not CodeView
 
    If CodeView Then
    
        ObjectSelected = iCodeView
        Toolbar1.Buttons(13).Image = ImageList1.ListImages(10).Index
        Toolbar1.Buttons(13).ToolTipText = "Form Designer..."
        Toolbar1.Buttons(7).Enabled = EnablePaste
        Toolbar1.Buttons(10).Enabled = False
        mnucut.Enabled = frmCode.EnableCutCopy
        mnudelete.Enabled = frmCode.EnableCutCopy
        mnupaste.Enabled = EnablePaste
        mnuall.Enabled = True
        mnufront.Enabled = False
        mnuBack.Enabled = False
        mnufind.Enabled = True
        mnugoto.Enabled = True
        mnureplace.Enabled = True
        mnuconvert.Enabled = True
        mnugrid.Enabled = False
        mnuinsert.Enabled = True
        DevButton(1).Enabled = False
        DevButton_DevButtonMouseUp 0, 0, 0, 0, 0
    Else
        ObjectSelected = iDialogView
        Toolbar1.Buttons(13).Image = ImageList1.ListImages(11).Index
        Toolbar1.Buttons(13).ToolTipText = "View Code..."
        Toolbar1.Buttons(10).Enabled = True
        Toolbar1.Buttons(7).Enabled = False ' disable paste button
        Toolbar1.Buttons(15).Enabled = False
        Toolbar1.Buttons(16).Enabled = False
        mnucut.Enabled = False
        mnucopy.Enabled = False
        mnudelete.Enabled = False
        mnupaste.Enabled = False ' disable paste menu item
        mnuall.Enabled = False
        mnufind.Enabled = False
        mnugoto.Enabled = False
        mnureplace.Enabled = False
        mnuconvert.Enabled = False
        mnugrid.Enabled = True
        mnuinsert.Enabled = False
        DevButton(1).Enabled = True
        DevButton_DevButtonMouseUp 1, 0, 0, 0, 0
        DevButton(1).ForeColor = vbBlack
    End If
    
End Sub

Private Sub PicHideBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoHideButtonState 1
End Sub

Private Sub PicHideBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoHideButtonState 0
End Sub

Private Sub PicHideBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoHideButtonState 0
    ShowToolBox False ' hide the toolbox
End Sub

Private Sub PicShow1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If FirstTimeLoad Then Exit Sub
    PicShow1.BorderStyle = 1
End Sub

Private Sub PicShow1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If FirstTimeLoad Then Exit Sub
    PicShow1.BorderStyle = 0
    ShowToolBox True ' show the toolbox
End Sub

Private Sub PicTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoHideButtonState 2
End Sub

Private Sub PicToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' MDIForm_MouseMove 1, 0, 0, 0
   
    PicTitleBar_MouseMove 1, 0, 0, 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

    sErrorText = ""
    
    Select Case UCase(Button.Key)
        Case "T_NEW"    ' Call Menu New
            mnunew_Click
        Case "T_OPEN"   ' Call Menu open
            mnuopen_Click
        Case "T_SAVE"
            mnusave_Click ' Call Menu Save
        Case "T_CUT"
            mnucut_Click ' Call menu Cut
        Case "T_COPY"
            mnucopy_Click ' Call Menu Copy
        Case "T_FIND"
            mnufind_Click
        Case "T_PASTE"
            mnupaste_Click ' Call menu Paste
        Case "T_BACK"
            mnuback_Click
        Case "T_FONT"
            mnufront_Click
        Case "T_FORM"
            mnuView_Click

        Case "T_RUN"
            Toolbar1.Buttons(1).Enabled = False
            Toolbar1.Buttons(2).Enabled = False
            Toolbar1.Buttons(3).Enabled = False
            Toolbar1.Buttons(13).Enabled = False ' disable code view button
            Toolbar1.Buttons(10).Enabled = False ' disable the run button
            Toolbar1.Buttons(11).Enabled = True ' enable the stop button
            Toolbar1.Buttons(15).Enabled = False
            Toolbar1.Buttons(16).Enabled = False
            mnunew.Enabled = False
            mnuopen.Enabled = False
            mnusave.Enabled = False
            mnumake.Enabled = False
            mnufront.Enabled = False
            mnugrid.Enabled = False
            mnuBack.Enabled = False
            mnucut.Enabled = False
            mnuView.Enabled = False
            mnufind.Enabled = False
            mnugoto.Enabled = False
            mnureplace.Enabled = False
            DlgDesignTime = False ' we'r in run time mode
            PicToolBar.Visible = False ' hide the toolbox
            mnutoolbox.Enabled = False
            
            DialogRun frmWorkArea
            ' Remmber the forms data
            RemberFormData frmWorkArea
            ' setup the script control
            SetupScript
            ' run the main code here
            dScript.RunCode frmCode.txtCode.Text
            
            If dScript.HasError Then
                GetLastScriptError
            End If
            
        Case "T_STOP"
            IdeStop
    End Select
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicTitleBar_MouseMove 1, 0, 0, 0
End Sub

Private Sub Tray1_MouseUp(Button As Integer)
    Me.WindowState = 0
    Tray1.Visible = False
    Tray1.ToolTip = ""
    MDIForm1.Visible = True
End Sub
