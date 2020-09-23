VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test Example - Plug-in"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   450
      Left            =   4215
      TabIndex        =   3
      Top             =   2955
      Width           =   1095
   End
   Begin VB.TextBox txtCode 
      Height          =   2670
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   135
      Width           =   5625
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show IDE Grid"
      Height          =   450
      Left            =   2325
      TabIndex        =   1
      Top             =   2955
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Editor Text"
      Height          =   435
      Left            =   150
      TabIndex        =   0
      Top             =   2955
      Width           =   1980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Editor As Object ' Code Editor object
Public DevIDE As Object ' IDE object

Private Sub Command1_Click()
    txtCode.Text = Editor.Text ' get the code editors text
End Sub

Private Sub Command2_Click()
    DevIDE.mnugrid_Click ' Call the mnugrid_Click in the IDE
End Sub

Private Sub Command3_Click()
    Unload Form1 ' unloads this form
End Sub

