VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FrmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pictab 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   1
      Left            =   195
      ScaleHeight     =   1935
      ScaleWidth      =   6000
      TabIndex        =   12
      Top             =   1980
      Width           =   6000
      Begin SHDocVwCtl.WebBrowser sWebB 
         Height          =   1920
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   6000
         ExtentX         =   10583
         ExtentY         =   3387
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.PictureBox pictab 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Index           =   2
      Left            =   225
      ScaleHeight     =   1545
      ScaleWidth      =   5820
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   5820
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Might you have any problems or questions about this program please inform me direct at the address provided below:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         TabIndex        =   10
         Top             =   105
         Width           =   5625
      End
      Begin VB.Label lblmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vbdream2k@yahoo.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1785
         TabIndex        =   9
         Top             =   690
         Width           =   2085
      End
   End
   Begin VB.PictureBox pictab 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Index           =   0
      Left            =   330
      ScaleHeight     =   1545
      ScaleWidth      =   5820
      TabIndex        =   5
      Top             =   5145
      Visible         =   0   'False
      Width           =   5820
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Written by Ben Jones"
         Height          =   195
         Left            =   3990
         TabIndex        =   11
         Top             =   1185
         Width           =   1605
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "This program is freeware no parts may be used in commercial use applications without written permission."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   165
         TabIndex        =   7
         Top             =   615
         Width           =   5085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DM Tiny Develop is a free small scripting environment for building robust applications for your business clients."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   165
         TabIndex        =   6
         Top             =   75
         Width           =   5625
      End
   End
   Begin MSComctlLib.TabStrip sTab 
      Height          =   2370
      Left            =   135
      TabIndex        =   4
      Top             =   1605
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   4180
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Change Log"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contact"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   90
      ScaleHeight     =   1320
      ScaleWidth      =   6150
      TabIndex        =   1
      Top             =   90
      Width           =   6150
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   45
         Picture         =   "FrmAbout.frx":0000
         Top             =   75
         Width           =   3870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   3
         Top             =   900
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Produced for Microsoft Windows Systems."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   2
         Top             =   330
         Width           =   2700
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5025
      TabIndex        =   0
      Top             =   4140
      Width           =   1215
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004 Ben Jones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   210
      TabIndex        =   14
      Top             =   4200
      Width           =   2070
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MoveTab(Index As Integer)
Dim I As Integer
    For I = 0 To pictab.Count - 1
        pictab(I).Visible = False
    Next
    I = 0
    pictab(Index).Top = 1980
    pictab(Index).Left = 195
    pictab(Index).Visible = True
End Sub

Private Sub Command1_Click()
    Unload FrmAbout
End Sub

Private Sub Form_Load()
    FrmAbout.Icon = Nothing
    If Not IsFileHere(ApplicationPath & "doc\ChangeLog.html") Then
        MsgBox "Unable to find " & ApplicationPath & "doc\ChangeLog.html", vbCritical, "File not found"
        sWebB.Navigate "about:blank"
    Else
        sWebB.Navigate ApplicationPath & "doc\ChangeLog.html"
    End If
    sTab_Click
End Sub

Private Sub lblmail_Click()
Dim iVal As Long
    iVal = ShellExecute(FrmAbout.hwnd, vbNullString, "mailto:" & lblmail.Caption, vbNullString, vbNullString, 3)
End Sub

Private Sub sTab_Click()
    MoveTab sTab.SelectedItem.Index - 1
End Sub
