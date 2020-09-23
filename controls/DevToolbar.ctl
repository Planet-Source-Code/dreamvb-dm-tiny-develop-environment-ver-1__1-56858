VERSION 5.00
Begin VB.UserControl DevToolbar 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   ScaleHeight     =   104
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   189
   Begin VB.PictureBox PicButtons 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   105
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   4
      Top             =   1050
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicBack 
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   0
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.VScrollBar vBar 
         Height          =   510
         Left            =   2640
         TabIndex        =   2
         Top             =   90
         Width           =   225
      End
      Begin VB.PictureBox PicBuffer 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   -15
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   3
         Top             =   15
         Width           =   2610
      End
   End
   Begin VB.PictureBox PicButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   170
      TabIndex        =   0
      Top             =   990
      Visible         =   0   'False
      Width           =   2550
   End
End
Attribute VB_Name = "DevToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private Const ButImgSize As Integer = 16
Private Const ButImgXPos As Integer = 2.5
Private Const ButImgYPos As Integer = 2

Private Const ButtonHeight As Integer = 20

Private Type TButton
    ButtonCaption As String
    ButtonImgIndex As Integer
    ButtonKey As String
End Type

Private Type DevToolbar
    Count As Integer
    Button() As TButton
End Type

Private ButtonCounter As Integer
Dim ButtonYPos As Integer
Dim DevButtonBar As DevToolbar
Dim ButtonIndex As Integer

Event DevToolBarMouseDown(Button As Integer, Index As Integer, Key As String)
Event DevToolBarMouseUp(Button As Integer, Index As Integer, Key As String)
Event DevToolBarMouseMove(Button As Integer, Index As Integer, Key As String)

Public Sub HideFocus()
On Error Resume Next
    If DevButtonBar.Count = -1 Then Exit Sub
    MakeSingleButton DevButtonBar.Button(ButtonIndex).ButtonCaption, DevButtonBar.Button(ButtonIndex).ButtonImgIndex, 0
    BitBlt PicBuffer.hdc, 0, ButtonIndex * ButtonHeight, PicButton.Width, ButtonHeight, PicButton.hdc, 0, 0, vbSrcCopy
    PicBuffer.Refresh
End Sub

Public Sub DrawToolBar()
Dim I As Integer
On Error Resume Next

    If DevButtonBar.Count = -1 Then Exit Sub
    PicBuffer.Height = (DevButtonBar.Count * ButtonHeight) + ButtonHeight
    If PicBuffer.Height = 1 Then PicBuffer.Height = 20
    
    For I = 0 To DevButtonBar.Count
        ButtonYPos = (I * ButtonHeight)
        MakeSingleButton DevButtonBar.Button(I).ButtonCaption, DevButtonBar.Button(I).ButtonImgIndex, 0
        BitBlt PicBuffer.hdc, 0, ButtonYPos, PicButton.Width, ButtonHeight, PicButton.hdc, 0, 0, vbSrcCopy
        PicBuffer.Refresh
    Next
    
    vBar.Max = (DevButtonBar.Count * ButtonHeight) - picBack.Height + 24
    vBar.Visible = vBar.Max > 0
    
End Sub

Private Sub PicBuffer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    DoEvents
    If DevButtonBar.Count < 0 Then Exit Sub
    ButtonIndex = Fix(y / ButtonHeight)
    If ButtonIndex <= 0 Then ButtonIndex = 0
    MakeSingleButton DevButtonBar.Button(ButtonIndex).ButtonCaption, DevButtonBar.Button(ButtonIndex).ButtonImgIndex, 2
    BitBlt PicBuffer.hdc, 0, ButtonIndex * ButtonHeight, PicButton.Width, ButtonHeight, PicButton.hdc, 0, 0, vbSrcCopy
    PicBuffer.Refresh
    RaiseEvent DevToolBarMouseDown(Button, ButtonIndex, DevButtonBar.Button(ButtonIndex).ButtonKey)
End Sub

Private Sub PicBuffer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim kPos As Integer
On Error Resume Next
    'If DevButtonBar.Count < 0 Then Exit Sub
    MakeSingleButton DevButtonBar.Button(ButtonIndex).ButtonCaption, DevButtonBar.Button(ButtonIndex).ButtonImgIndex, 0
    BitBlt PicBuffer.hdc, 0, ButtonIndex * ButtonHeight, PicButton.Width, ButtonHeight, PicButton.hdc, 0, 0, vbSrcCopy
    PicBuffer.Refresh
    ButtonIndex = Fix(y / ButtonHeight)
    kPos = Fix(y / ButtonHeight)
    MakeSingleButton DevButtonBar.Button(kPos).ButtonCaption, DevButtonBar.Button(kPos).ButtonImgIndex, 1
    BitBlt PicBuffer.hdc, 0, kPos * ButtonHeight, PicButton.Width, ButtonHeight, PicButton.hdc, 0, 0, vbSrcCopy
    PicBuffer.Refresh
    RaiseEvent DevToolBarMouseMove(Button, ButtonIndex, DevButtonBar.Button(ButtonIndex).ButtonKey)
    
End Sub

Private Sub PicBuffer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If DevButtonBar.Count < 0 Then Exit Sub
    If ButtonIndex > DevButtonBar.Count Then ButtonIndex = DevButtonBar.Count - 1
    If ButtonIndex <= -1 Then ButtonIndex = 0
    MakeSingleButton DevButtonBar.Button(ButtonIndex).ButtonCaption, DevButtonBar.Button(ButtonIndex).ButtonImgIndex, 1
    BitBlt PicBuffer.hdc, 0, ButtonIndex * ButtonHeight, PicButton.Width, ButtonHeight, PicButton.hdc, 0, 0, vbSrcCopy
    PicBuffer.Refresh
    RaiseEvent DevToolBarMouseUp(Button, ButtonIndex, DevButtonBar.Button(ButtonIndex).ButtonKey)
End Sub

Private Sub vBar_Change()
    PicBuffer.Top = -vBar.Value
End Sub

Private Sub vBar_GotFocus()
    HideFocus
End Sub

Private Sub vBar_Scroll()
    HideFocus
    vBar_Change
End Sub

Public Sub AddButton(Optional Caption As String = "Button", Optional ButImgIndex As Integer = 0, Optional ButKey As String = " ")
On Error Resume Next
    DevButtonBar.Count = DevButtonBar.Count + 1
    ReDim Preserve DevButtonBar.Button(DevButtonBar.Count)
    DevButtonBar.Button(DevButtonBar.Count).ButtonCaption = Caption
    DevButtonBar.Button(DevButtonBar.Count).ButtonImgIndex = ButImgIndex
    DevButtonBar.Button(DevButtonBar.Count).ButtonKey = ButKey
    SetupControlBar
End Sub

Private Sub DrawEffect(Direction As Integer)
    If Direction = 1 Then
        PicButton.Line (PicButton.ScaleWidth, 0)-(0, 0), vbWhite
        PicButton.Line (0, PicButton.ScaleHeight)-(0, -1), vbWhite
        PicButton.Line (PicButton.ScaleWidth - 1, 0)-(PicButton.ScaleWidth - 1, PicButton.ScaleHeight), vbApplicationWorkspace
        PicButton.Line (PicButton.ScaleWidth, PicButton.ScaleHeight - 1)-(-1, PicButton.ScaleHeight - 1), vbApplicationWorkspace
        Exit Sub
    ElseIf Direction = 2 Then
        PicButton.Line (PicButton.ScaleWidth, 0)-(0, 0), vbApplicationWorkspace
        PicButton.Line (0, PicButton.ScaleHeight)-(0, -1), vbApplicationWorkspace
        PicButton.Line (PicButton.ScaleWidth - 1, 0)-(PicButton.ScaleWidth - 1, PicButton.ScaleHeight), vbWhite
        PicButton.Line (PicButton.ScaleWidth, PicButton.ScaleHeight - 1)-(-1, PicButton.ScaleHeight - 1), vbWhite
    Else
        PicButton.Line (PicButton.ScaleWidth, 0)-(0, 0), vbButtonFace
        PicButton.Line (0, PicButton.ScaleHeight)-(0, -1), vbButtonFace
        PicButton.Line (PicButton.ScaleWidth - 1, 0)-(PicButton.ScaleWidth - 1, PicButton.ScaleHeight), vbButtonFace
        PicButton.Line (PicButton.ScaleWidth, PicButton.ScaleHeight - 1)-(-1, PicButton.ScaleHeight - 1), vbButtonFace
    End If
End Sub

Private Sub MakeSingleButton(Optional Caption As String = "Button", Optional ButtonImgIndex As Integer = 0, Optional ButtonStyle As Integer)
    DoEvents
    PicButton.Cls
    TransparentBlt PicButton.hdc, ButImgXPos, ButImgYPos, ButImgSize, ButImgSize, PicButtons.hdc, ButImgSize * ButtonImgIndex, 0, ButImgSize, ButImgSize, RGB(255, 0, 255)
    PicButton.CurrentX = 20
    PicButton.CurrentY = 3
    PicButton.Print Caption
    DrawEffect ButtonStyle
    PicButton.Refresh
End Sub

Public Sub ResetButton()
On Error Resume Next
    PicBuffer.Cls
    PicButton.Cls
    vBar.Value = 0
    DevButtonBar.Count = -1
    Erase DevButtonBar.Button()
End Sub

Public Sub SetupControlBar()
    
    vBar.Top = 0: vBar.Left = (picBack.ScaleWidth - vBar.Width)
    vBar.Height = (picBack.ScaleHeight)
    PicBuffer.Left = 0: PicBuffer.Top = 0: PicBuffer.Width = (picBack.ScaleWidth)
    If ((DevButtonBar.Count * ButtonHeight) - picBack.Height) > 0 Then
       PicButton.Width = (vBar.Left)
    Else
        PicButton.Width = PicBuffer.ScaleWidth
    End If
    
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    picBack.Height = (UserControl.ScaleHeight)
    picBack.Width = (UserControl.ScaleWidth)
    SetupControlBar
    If Err Then UserControl.Size 90, 90
End Sub
Public Function ButtonCaption(Index As Integer) As String
    ButtonCaption = DevButtonBar.Button(Index).ButtonCaption
End Function

Public Function ButtonKey(Index As Integer) As String
    ButtonKey = DevButtonBar.Button(Index).ButtonKey
End Function
Public Function ButtonImgIndex(Index As Integer) As Integer
    ButtonImgIndex = DevButtonBar.Button(Index).ButtonImgIndex
End Function

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = PicButtons.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set PicButtons.Picture = New_Picture
    PropertyChanged "Picture"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub
