VERSION 5.00
Begin VB.UserControl DevButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   164
   ToolboxBitmap   =   "DevButton.ctx":0000
End
Attribute VB_Name = "DevButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_Caption As String
Const m_def_Caption = "Button"

Event DevButtonMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event DevButtonMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event DevButtonMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Private Sub DrawCaption(ButtonState As Integer, Optional nEnabled As Boolean)
If UserControl.Enabled Then
    UserControl.Cls
    UserControl.CurrentX = 6
    UserControl.CurrentY = (UserControl.ScaleHeight - TextHeight(m_Caption)) / 2
    UserControl.Print m_Caption
    DrawEffect ButtonState
Else
    UserControl.ForeColor = vbWhite
    UserControl.Cls
    UserControl.CurrentX = 6
    UserControl.CurrentY = (UserControl.ScaleHeight - TextHeight(m_Caption)) / 2
    UserControl.Print m_Caption
    DrawEffect ButtonState
    UserControl.ForeColor = &H808080
    UserControl.CurrentX = 5
    UserControl.CurrentY = (UserControl.ScaleHeight - TextHeight(m_Caption)) / 2 + 1
    UserControl.Print m_Caption
    DrawEffect ButtonState
End If
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawCaption 0
    RaiseEvent DevButtonMouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent DevButtonMouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawCaption 1
    RaiseEvent DevButtonMouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
    DrawCaption 1
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawCaption 1
End Property

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Caption = m_def_Caption
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
End Sub

Private Sub UserControl_Show()
    DrawCaption 1
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, vbBlack)
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawCaption 1
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawCaption 1
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    DrawCaption 1, New_Enabled
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    DrawCaption 1, UserControl.Enabled
End Property

