VERSION 5.00
Begin VB.Form frmWorkArea 
   BackColor       =   &H80000009&
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   Icon            =   "frmWorkArea.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4320
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   0
      ItemData        =   "frmWorkArea.frx":08CA
      Left            =   1290
      List            =   "frmWorkArea.frx":08CC
      TabIndex        =   4
      Top             =   1890
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtA 
      Height          =   525
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   1890
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox PicImg 
      Height          =   495
      Index           =   0
      Left            =   105
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   930
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdBut 
      Height          =   495
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   1590
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmWorkArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private tForeCol As Long

Public Function AppPath() As String
Dim lzPath As String
    lzPath = FixPath(App.Path)
    AppPath = lzPath
    lzPath = ""
End Function

Public Sub TSavePicture(PicObject As PictureBox, SaveFileName)
    SavePicture PicObject.Image, SaveFileName
End Sub

Public Function DrawLine(mObject As Object, X1, Y1, X2, Y2, mColor)
    mObject.Line (X1, Y1)-(X2, Y2), mColor
End Function

Public Function GetPoint(mObject As Object, x, Y) As Long
    GetPoint = mObject.Point(x, Y)
End Function

Public Sub Plot(mObject As Object, x, Y, bColor)
    mObject.PSet (x, Y), bColor
End Sub
Public Sub CenterDialog()
    frmWorkArea.Left = (Screen.Width - frmWorkArea.Width) / 2
    frmWorkArea.Top = (Screen.Height - frmWorkArea.Height) / 2
End Sub

Public Property Get ForeColorf() As Long
    ForeColorf = tForeCol
End Property

Public Property Let ForeColorf(ByVal vNewValue As Long)
    tForeCol = vNewValue
End Property

Public Sub UnloadDialog(): End: End Sub

Private Sub CmdBut_Click(Index As Integer)
    dScript.RunCode2 "call CmdBut" & Index & "_Click()"
End Sub

Private Sub Form_Initialize()
    Dim x As Long
    x = InitCommonControls
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call Dialog_MouseMove(" & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub lblA_Click(Index As Integer)
    dScript.RunCode2 "call lblA" & Index & "_Click()"
End Sub

Private Sub lblA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call lblA" & Index & "_MouseDown(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub lblA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call lblA" & Index & "_MouseMove(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub lblA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call lblA" & Index & "_MouseUp(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub lstA_Click(Index As Integer)
    dScript.RunCode2 "call lstA" & Index & "_Click()"
End Sub

Private Sub lstA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call lstA" & Index & "_MouseDown(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub lstA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call lstA" & Index & "_MouseMove(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub lstA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call lstA" & Index & "_MouseUp(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub PicImg_Click(Index As Integer)
    dScript.RunCode2 "PicImg" & Index & "_Click()"
End Sub

Private Sub PicImg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call PicImg" & Index & "_MouseDown(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub PicImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call PicImg" & Index & "_MouseMove(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub PicImg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call PicImg" & Index & "_MouseUp(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub txtA_Change(Index As Integer)
    dScript.RunCode2 "call txtA" & Index & "_Change(index)"
End Sub

Private Sub txtA_Click(Index As Integer)
    dScript.RunCode2 "call txtA" & Index & "_Click(index)"
End Sub

Private Sub txtA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call txtA" & Index & "_MouseDown(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub txtA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call txtA" & Index & "_MouseMove(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub

Private Sub txtA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    dScript.RunCode2 "call txtA" & Index & "_MouseUp(" & Index & "," & Button & "," & Shift & "," & x & "," & Y & ")"
End Sub
