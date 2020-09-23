VERSION 5.00
Begin VB.Form frmCode 
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   5475
   WindowState     =   2  'Maximized
   Begin Project1.devToolTip devToolTip1 
      Height          =   240
      Left            =   2595
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   423
      Caption         =   ""
   End
   Begin VB.PictureBox PicMarginBar 
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   375
      Width           =   345
      Begin VB.PictureBox PicLine 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   330
         ScaleHeight     =   300
         ScaleWidth      =   15
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.ComboBox CboList 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   2565
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   345
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   2430
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CboTmp As String
Dim sTextLineA As String

Private Sub UpDateInSight(sFuncName As String)
Dim tPoint As POINTAPI, aStr As String

    GetCaretPos tPoint
    devToolTip1.Visible = True
    devToolTip1.Left = (tPoint.X * Screen.TwipsPerPixelX) + tPoint.X
    devToolTip1.Top = (tPoint.Y * Screen.TwipsPerPixelY) + 680
    aStr = " " & sFuncName & " "
    aStr = Replace(aStr, "_", vbCrLf & "  ")
    devToolTip1.Caption = aStr
    aStr = ""
    
End Sub
Public Function EnableCutCopy() As Boolean
    If Len(txtCode.SelText) = 0 Then EnableCutCopy = False: Exit Function
    EnableCutCopy = True
End Function

Private Sub CboList_Change()
    CboList.Text = CboTmp
End Sub

Private Sub CboList_Click()
Dim lpos As Long, hPos As Long
    CboTmp = CboList.Text
    
    lpos = InStr(1, txtCode.Text, LTrim(CboTmp), vbTextCompare)
    hPos = InStr(lpos + 1, txtCode.Text, vbCrLf)
    If (lpos > 0) And (hPos > 0) Then
        txtCode.SelStart = (hPos + 1)
        txtCode.SetFocus
    End If
    
    lpos = 0: hPos = 0
    
End Sub

Public Sub Form_Activate()
    MDIForm1.Toolbar1.Buttons(5).Enabled = EnableCutCopy
    MDIForm1.Toolbar1.Buttons(6).Enabled = EnableCutCopy
    lstVBFunctions txtCode.Text, CboList
    RemoveControlButton frmCode, 6, False
    PicMarginBar.Visible = AppConfig.EditMarginBar
    
    If AppConfig.EditMarginBar Then
        txtCode.Left = 345
    Else
        PicMarginBar.Visible = False
        txtCode.Left = 0
    End If
    
    txtCode.FontName = AppConfig.EditFont
    txtCode.FontSize = AppConfig.EditFontSize
    txtCode.ForeColor = AppConfig.EditFontColor
    txtCode.BackColor = AppConfig.EditFontBackColor
    SetMargin 8, txtCode
    Form_Resize
    
End Sub

Private Sub Form_Load()
    FlatBorder txtCode.hwnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CboTmp = ""
End Sub

Private Sub Form_Resize()
On Error Resume Next
    txtCode.Width = (frmCode.ScaleWidth - txtCode.Left)
    txtCode.Height = (frmCode.ScaleHeight - CboList.Height - 60)
    If AppConfig.EditMarginBar Then PicMarginBar.Height = (frmCode.ScaleHeight - PicMarginBar.Top - 20): PicLine.Height = PicMarginBar.Height
End Sub

Private Sub txtCode_Change()
    MDIForm1.lblTextPos.Caption = "Ln " & GetCurrentLineNumber(txtCode) & ", " & "Col " & GetColumn(txtCode)
    Modified = True
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then txtCode.SelText = Space(AppConfig.EditTabSize): KeyAscii = 0
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Dim sBuf As String, iPos As Long, sFunction As String, vLst As Variant
Dim FuncCount As Long, I As Long

    txtCode_Change
    
    If KeyCode = 13 Then
        lstVBFunctions txtCode.Text, CboList
        devToolTip1.Visible = False
    End If

    If (KeyCode = 16) And AppConfig.EditCodeinsight Then
    
        txtCode_MouseDown 1, 0, 0, 0
        sBuf = sTextLineA
        
        For I = 1 To Len(sBuf)
            C = Mid(sBuf, I, 1)
            If C = " " Then iPos = I
        Next
        
        C = ""
        I = 0

        If iPos = 0 Then
            sFunction = sBuf
        Else
            sFunction = Mid(sBuf, iPos + 1, Len(sBuf))
        End If

        sBuf = ""
        iPos = 0
        
        vLst = Split(sFunction, "(")

        FuncCount = UBound(vLst) - 1
        sFunction = vLst(FuncCount)

        If ItemExists(sFunction) And Len(sFunction) > 0 Then
            UpDateInSight GetItemKey(sFunction)
        Else
            devToolTip1.Visible = False
        End If
        
        Erase vLst
        sTextLineA = ""
        sFunction = ""
        sBuf = ""
    End If

    If (KeyCode = (Shift + vbKey0)) Then devToolTip1.Visible = False

End Sub

Private Sub txtCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Button = 2 Then sTextLineA = GetLineText(txtCode)
End Sub

Private Sub txtCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MDIForm1.MDIForm_MouseMove 1, 0, 0, 0
End Sub

Private Sub txtCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MDIForm1.Toolbar1.Buttons(5).Enabled = EnableCutCopy
    MDIForm1.Toolbar1.Buttons(6).Enabled = EnableCutCopy
    MDIForm1.mnucut.Enabled = EnableCutCopy
    MDIForm1.mnucopy.Enabled = EnableCutCopy
    MDIForm1.mnudelete.Enabled = EnableCutCopy
    MDIForm1.mnuupper.Enabled = EnableCutCopy
    MDIForm1.mnulower.Enabled = EnableCutCopy
    MDIForm1.mnuinvert.Enabled = EnableCutCopy
    MDIForm1.mnuspacetab.Enabled = EnableCutCopy
    MDIForm1.mnutpsapace.Enabled = EnableCutCopy
    lstVBFunctions txtCode.Text, CboList
    devToolTip1.Visible = False
    txtCode_Change
End Sub
