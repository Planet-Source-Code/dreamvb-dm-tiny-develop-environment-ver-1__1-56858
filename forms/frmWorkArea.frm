VERSION 5.00
Begin VB.Form frmWorkArea 
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3285
   ScaleWidth      =   4665
   Begin Project1.devToolTip devToolTip1 
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   2910
      Visible         =   0   'False
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   423
      Caption         =   " "
   End
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   0
      Left            =   1380
      TabIndex        =   5
      Top             =   1905
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox txtA 
      Height          =   525
      Index           =   0
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1890
      Visible         =   0   'False
      Width           =   1170
   End
   Begin Project1.bSelect Hangle 
      Height          =   90
      Left            =   150
      TabIndex        =   3
      Top             =   2595
      Visible         =   0   'False
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   159
      MousePointer    =   8
   End
   Begin VB.PictureBox PicImg 
      Height          =   495
      Index           =   0
      Left            =   75
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   930
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdBut 
      Caption         =   "Button"
      Height          =   400
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   1024
   End
   Begin VB.Shape Selection 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   390
      Left            =   390
      Top             =   2505
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      Height          =   495
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   1590
      Visible         =   0   'False
      Width           =   1215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmWorkArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Oldx As Integer, OldY As Integer, tForeCol As Long
Dim CanMove As Boolean, CanResize As Boolean, nTmp As String

Public Function AppPath() As String
Dim lzPath As String
    lzPath = ProjectFolder
    AppPath = lzPath
    lzPath = ""
End Function

Public Sub TSavePicture(PicObject As PictureBox, SaveFileName)
    SavePicture PicObject.Image, SaveFileName
End Sub

Public Function DrawLine(mObject, X1, Y1, X2, Y2, mColor)
    mObject.Line (X1, Y1)-(X2, Y2), mColor
End Function

Public Function GetPoint(mObject As Object, X, Y) As Long
    GetPoint = mObject.Point(X, Y)
End Function

Public Sub Plot(mObject As Object, X, Y, bColor)
    mObject.PSet (X, Y), bColor
End Sub

Private Sub UpdateInfoTip(TCtrName As Object, Button As Integer)
    If Not AppConfig.ShowControlHints Then Exit Sub
    devToolTip1.Visible = (Button = 0)
    devToolTip1.Left = TCtrName.Left
    devToolTip1.Top = (TCtrName.Top + TCtrName.Height + 120)
    devToolTip1.Caption = " " & TypeName(TCtrName) & ": " & TCtrName.Name & "(" & TCtrName.Index & ") " _
    & "  " & vbCrLf & "  Origin: " & TCtrName.Left & ", " & TCtrName.Top _
    & "  " & vbCrLf & "  Size: " & TCtrName.Width & " x " & TCtrName.Height
    devToolTip1.ZOrder vbBringToFront
End Sub
Public Sub CenterDialog()
    If Not inIde Then
        frmWorkArea.Left = (MDIForm1.ScaleWidth - frmWorkArea.Width) / 2
        frmWorkArea.Top = (MDIForm1.ScaleHeight - frmWorkArea.Height) / 2
    End If
End Sub
Public Property Get ForeColorf() As Long
    ForeColorf = tForeCol
End Property

Public Property Let ForeColorf(ByVal vNewValue As Long)
    tForeCol = vNewValue
End Property

Public Sub UnloadDialog()
    MsgBox "Please use the stop button in the ide", vbInformation, "Stop"
End Sub
Public Sub HideSelection()
    Selection.Visible = False
    Hangle.Visible = False
    MDIForm1.Toolbar1.Buttons(5).Enabled = False
End Sub

Public Sub MakeSelection(mShow As Boolean)
    Hangle.ZOrder 0
    Selection.Top = (TheObjectName.Top - 65)
    Selection.Left = (TheObjectName.Left - 65)
    Selection.Width = (TheObjectName.Width + 130)
    Selection.Height = (TheObjectName.Height + 130)
    Hangle.Top = (TheObjectName.Top + Selection.Height)
    Hangle.Left = (TheObjectName.Left + Selection.Width)
    Selection.Visible = mShow
    Hangle.Visible = mShow
    MDIForm1.lblPosition.Caption = TheObjectName.Top & ", " & TheObjectName.Left
    MDIForm1.lblSize.Caption = TheObjectName.Width & ", " & TheObjectName.Height
    Modified = True ' chnages have been make
End Sub

Public Sub TMouseUP(mObject As Object, Button As Integer, X As Single, Y As Single)
    If Not inIde Then Exit Sub
    If mObject.Top <= 0 Then mObject.Top = 0
    If mObject.Left <= 0 Then mObject.Left = 0
    
    If (mObject.Top + mObject.Height) >= frmWorkArea.Height Then mObject.Top = (frmWorkArea.Height - mObject.Height * 2 + 120)
    If (mObject.Left + mObject.Width) >= frmWorkArea.Width Then mObject.Left = (frmWorkArea.Width - mObject.Width - 120)
    MakeSelection True
    CanMove = False
    ObjectSelected = iDialogView ' dialog object selected
    MDIForm1.Toolbar1.Buttons(5).Enabled = True
    MDIForm1.Toolbar1.Buttons(15).Enabled = True
    MDIForm1.Toolbar1.Buttons(16).Enabled = True
    MDIForm1.mnucut.Enabled = True
    MDIForm1.mnudelete.Enabled = True
    MDIForm1.mnufront.Enabled = True
    MDIForm1.mnuback.Enabled = True
End Sub
Public Sub TMouseDown(mObject As Object, Button As Integer, X As Single, Y As Single)
    If Not inIde Then Exit Sub
    Set TheObjectName = mObject
    mObject.ZOrder 0
    Oldx = X
    OldY = Y
    CanMove = True
    MakeSelection True
    If Button = vbRightButton Then PopupMenu frmmenu.Mnu1
End Sub
Private Sub MoveControl(mObject As Object, Button As Integer, X As Single, Y As Single)
    If Not inIde Then Exit Sub
    UpdateInfoTip mObject, Button ' show info tip on the control
    If Button = vbLeftButton Then
        mObject.Left = mObject.Left + (X - Oldx)
        mObject.Top = mObject.Top + (Y - OldY)
        MakeSelection True
        Modified = True
    End If
End Sub

Private Sub CmdBut_Click(Index As Integer)
    If inIde Then Exit Sub
    dScript.RunCode2 "call CmdBut" & Index & "_Click()"
End Sub

Private Sub CmdBut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown CmdBut(Index), Button, X, Y
End Sub

Private Sub CmdBut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl CmdBut(Index), Button, X, Y
End Sub

Private Sub CmdBut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP CmdBut(Index), Button, X, Y
End Sub

Private Sub Form_Load()
    inIde = True
    ObjectSelected = False
    RemoveControlButton frmWorkArea, 6, False
    frmWorkArea.Top = 120: frmWorkArea.Left = 120
    FormXPos = (frmWorkArea.Left / Screen.TwipsPerPixelX)
    FormYPos = (frmWorkArea.Top / Screen.TwipsPerPixelY)
    Set frmWorkArea.Icon = Nothing
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideSelection
    MDIForm1.mnucut.Enabled = False
    MDIForm1.mnufront.Enabled = False
    MDIForm1.mnuback.Enabled = False
    MDIForm1.mnudelete.Enabled = False
    MDIForm1.Toolbar1.Buttons(5).Enabled = False
    MDIForm1.Toolbar1.Buttons(15).Enabled = False
    MDIForm1.Toolbar1.Buttons(16).Enabled = False
    MDIForm1.lblPosition.Caption = "0, 0"
    MDIForm1.lblSize.Caption = "0, 0"
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    devToolTip1.Visible = False
    If inIde Then Exit Sub
    dScript.RunCode2 "call Dialog_MouseMove(" & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub Form_Resize()
    If Not inIde Then Exit Sub
    DrawGrid frmWorkArea, , MDIForm1.mnugrid.Checked
    If frmWorkArea.WindowState = 2 Then frmWorkArea.WindowState = 0
End Sub

Private Sub Hangle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CanResize = True
        Oldx = X
        OldY = Y
    End If
End Sub

Private Sub Hangle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If (Button = vbLeftButton And CanResize) Then
        Hangle.Top = Hangle.Top - (OldY - Y)
        Hangle.Left = Hangle.Left - (Oldx - X)
        Selection.Width = Hangle.Left - (Selection.Left - 8)
        Selection.Height = Hangle.Top - (Selection.Top - 8)
        MDIForm1.lblSize.Caption = Selection.Width & ", " & Selection.Height
    End If
End Sub

Private Sub Hangle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    CanResize = False
    TheObjectName.Width = Selection.Width - 130
    TheObjectName.Height = Selection.Height - 130
    If TheObjectName.Width <= 90 Then TheObjectName.Width = 90
    If TheObjectName.Height <= 90 Then TheObjectName.Height = 90
End Sub

Private Sub lblA_Click(Index As Integer)
    If inIde Then Exit Sub
    dScript.RunCode2 "call lblA" & Index & "_Click()"
End Sub

Private Sub lblA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown lblA(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call lblA" & Index & "_MouseDown(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub lblA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl lblA(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call lblA" & Index & "_MouseMove(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub lblA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP lblA(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call lblA" & Index & "_MouseUp(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub lstA_Click(Index As Integer)
    If inIde Then Exit Sub
    dScript.RunCode2 "call lstA" & Index & "_Click()"
End Sub

Private Sub lstA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown lstA(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call lstA" & Index & "_MouseDown(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub lstA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl lstA(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call lstA" & Index & "_MouseMove(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub lstA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP lstA(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call lstA" & Index & "_MouseUp(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub PicImg_Click(Index As Integer)
    If inIde Then Exit Sub
    dScript.RunCode2 "PicImg" & Index & "_Click()"
End Sub

Private Sub PicImg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown PicImg(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call PicImg" & Index & "_MouseDown(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub PicImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl PicImg(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call PicImg" & Index & "_MouseMove(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub PicImg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP PicImg(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call PicImg" & Index & "_MouseUp(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub txtA_Change(Index As Integer)
    If inIde Then txtA(Index).Text = nTmp
    If inIde Then Exit Sub
    dScript.RunCode2 "call txtA" & Index & "_Change(index)"
End Sub

Private Sub txtA_Click(Index As Integer)
    If inIde Then nTmp = txtA(Index).Text
    If inIde Then Exit Sub
    dScript.RunCode2 "call txtA" & Index & "_Click(index)"
End Sub

Private Sub txtA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown txtA(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call txtA" & Index & "_MouseDown(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub txtA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl txtA(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call txtA" & Index & "_MouseMove(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub

Private Sub txtA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP txtA(Index), Button, X, Y
    If inIde Then Exit Sub
    dScript.RunCode2 "call txtA" & Index & "_MouseUp(" & Index & "," & Button & "," & Shift & "," & X & "," & Y & ")"
End Sub
