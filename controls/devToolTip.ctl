VERSION 5.00
Begin VB.UserControl devToolTip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   ScaleHeight     =   870
   ScaleWidth      =   3765
   Begin VB.Label lblTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   510
   End
End
Attribute VB_Name = "devToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' My little tooltip control not the best by far but hay it does it's job
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblTip.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblTip.Caption() = New_Caption
    PropertyChanged "Caption"
    UserControl_Resize
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblTip.Caption = PropBag.ReadProperty("Caption", "Label1")
End Sub

Private Sub UserControl_Resize()
    UserControl.Size lblTip.Width, lblTip.Height
    UserControl.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lblTip.Caption, "Label1")
End Sub

