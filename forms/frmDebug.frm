VERSION 5.00
Begin VB.Form frmDebug 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DM Tiny Develop Version 1"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4230
      TabIndex        =   3
      Top             =   1185
      Width           =   1215
   End
   Begin VB.CommandButton cmdDebug 
      Caption         =   "&Debug"
      Height          =   375
      Left            =   5580
      TabIndex        =   2
      Top             =   1185
      Width           =   1215
   End
   Begin VB.PictureBox p1 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1710
      Left            =   0
      ScaleHeight     =   1710
      ScaleWidth      =   750
      TabIndex        =   0
      Top             =   0
      Width           =   750
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   6
         X1              =   165
         X2              =   525
         Y1              =   180
         Y2              =   495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   6
         Height          =   480
         Left            =   105
         Shape           =   3  'Circle
         Top             =   105
         Width           =   480
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   6
         Height          =   480
         Left            =   135
         Shape           =   3  'Circle
         Top             =   135
         Width           =   480
      End
   End
   Begin VB.Label lblError 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   945
      TabIndex        =   1
      Top             =   90
      Width           =   45
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DrawLine()
    frmDebug.Line (p1.ScaleWidth, lblError.Height + 145)-(frmDebug.ScaleWidth, lblError.Height + 145), vbApplicationWorkspace
    frmDebug.Line (p1.ScaleWidth, lblError.Height + 165)-(frmDebug.ScaleWidth, lblError.Height + 165), vbWhite
    frmDebug.Refresh
End Sub
Private Sub cmdClose_Click()
    ButtonPressed = 0
    Unload frmDebug
End Sub

Private Sub cmdDebug_Click()
    ButtonPressed = 1
    Unload frmDebug
End Sub

Private Sub Form_Load()
    Beep
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDebug = Nothing
End Sub
