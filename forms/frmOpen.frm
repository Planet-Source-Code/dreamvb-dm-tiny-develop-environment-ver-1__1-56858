VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Environment Options"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   4350
      Width           =   1215
   End
   Begin VB.PictureBox PicTab1 
      BorderStyle     =   0  'None
      Height          =   3345
      Index           =   1
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   6585
      TabIndex        =   15
      Top             =   5280
      Visible         =   0   'False
      Width           =   6585
      Begin VB.Frame Frame2 
         Caption         =   "Editor"
         Height          =   3015
         Left            =   120
         TabIndex        =   16
         Top             =   150
         Width           =   6165
         Begin VB.Frame Frame5 
            Caption         =   "Sample"
            Height          =   900
            Left            =   3045
            TabIndex        =   32
            Top             =   1980
            Width           =   2820
            Begin VB.PictureBox p1 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   555
               Left            =   150
               ScaleHeight     =   555
               ScaleWidth      =   2535
               TabIndex        =   33
               Top             =   225
               Width           =   2535
               Begin VB.PictureBox PicMargin 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000008&
                  Height          =   585
                  Left            =   -15
                  ScaleHeight     =   555
                  ScaleWidth      =   270
                  TabIndex        =   34
                  Top             =   -15
                  Width           =   300
               End
               Begin VB.Label l5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   " If Err = cdlCancel Then Exit Sub"
                  Height          =   195
                  Index           =   1
                  Left            =   375
                  TabIndex        =   35
                  Top             =   165
                  Width           =   2310
               End
            End
         End
         Begin VB.CheckBox chkMargin 
            Caption         =   "Show Margin Indicator Bar"
            Height          =   225
            Left            =   3150
            TabIndex        =   31
            Top             =   1680
            Width           =   2445
         End
         Begin VB.CheckBox chkcode 
            Caption         =   "Show Quick Info Tooltips"
            Height          =   195
            Left            =   3150
            TabIndex        =   30
            Top             =   1335
            Width           =   2490
         End
         Begin VB.Frame Frame4 
            Caption         =   "Tab Indent:"
            Height          =   855
            Left            =   3165
            TabIndex        =   27
            Top             =   375
            Width           =   2520
            Begin VB.TextBox txtTabSize 
               Height          =   285
               Left            =   960
               TabIndex        =   29
               Top             =   360
               Width           =   690
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Tab Size:"
               Height          =   195
               Left            =   195
               TabIndex        =   28
               Top             =   405
               Width           =   675
            End
         End
         Begin VB.Frame Frame3 
            Height          =   885
            Left            =   135
            TabIndex        =   21
            Top             =   1905
            Width           =   2235
            Begin VB.PictureBox PicEdBkCol 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   1125
               ScaleHeight     =   195
               ScaleWidth      =   765
               TabIndex        =   25
               Top             =   465
               Width           =   825
            End
            Begin VB.PictureBox PicEdFCol 
               BackColor       =   &H00000000&
               Height          =   240
               Left            =   105
               ScaleHeight     =   180
               ScaleWidth      =   765
               TabIndex        =   23
               Top             =   465
               Width           =   825
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Back Colour:"
               Height          =   195
               Left            =   1125
               TabIndex        =   24
               Top             =   195
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fore Colour:"
               Height          =   195
               Left            =   105
               TabIndex        =   22
               Top             =   195
               Width           =   855
            End
         End
         Begin VB.ComboBox cbofontsize 
            Height          =   315
            Left            =   195
            TabIndex        =   20
            Top             =   1185
            Width           =   2175
         End
         Begin VB.ComboBox cboFont 
            Height          =   315
            Left            =   195
            TabIndex        =   17
            Top             =   525
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Editor Colour Properties:"
            Height          =   195
            Left            =   210
            TabIndex        =   26
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label l10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font Size:"
            Height          =   195
            Left            =   210
            TabIndex        =   19
            Top             =   960
            Width           =   705
         End
         Begin VB.Label l6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font:"
            Height          =   195
            Left            =   210
            TabIndex        =   18
            Top             =   285
            Width           =   360
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4290
      TabIndex        =   5
      Top             =   4350
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   240
      Top             =   4260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicTab1 
      BorderStyle     =   0  'None
      Height          =   3180
      Index           =   0
      Left            =   285
      ScaleHeight     =   3180
      ScaleWidth      =   6585
      TabIndex        =   8
      Top             =   690
      Visible         =   0   'False
      Width           =   6585
      Begin VB.CheckBox chkhints 
         Caption         =   "Show hints for controls"
         Height          =   270
         Left            =   2730
         TabIndex        =   4
         Top             =   750
         Width           =   3345
      End
      Begin VB.CheckBox chkFormMove 
         Caption         =   "Allow form to be moved at design time"
         Height          =   315
         Left            =   2715
         TabIndex        =   3
         Top             =   330
         Width           =   3435
      End
      Begin VB.Frame Frame1 
         Caption         =   "Grid"
         Height          =   2700
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   2415
         Begin VB.PictureBox picGridCol 
            BackColor       =   &H00000000&
            Height          =   300
            Left            =   180
            ScaleHeight     =   240
            ScaleWidth      =   1710
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1770
         End
         Begin VB.TextBox txtHeight 
            Height          =   285
            Left            =   1080
            TabIndex        =   2
            Top             =   1395
            Width           =   690
         End
         Begin VB.TextBox txtWidth 
            Height          =   285
            Left            =   1080
            TabIndex        =   1
            Top             =   990
            Width           =   690
         End
         Begin VB.CheckBox chkGrid 
            Caption         =   "Show Grid"
            Height          =   210
            Left            =   180
            TabIndex        =   0
            Top             =   255
            Width           =   1800
         End
         Begin VB.Label l4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grid Colour:"
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   1845
            Width           =   825
         End
         Begin VB.Label l3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height:"
            Height          =   195
            Left            =   405
            TabIndex        =   12
            Top             =   1455
            Width           =   510
         End
         Begin VB.Label l2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Width:"
            Height          =   195
            Left            =   405
            TabIndex        =   11
            Top             =   1035
            Width           =   465
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grid Size:"
            Height          =   195
            Left            =   225
            TabIndex        =   10
            Top             =   645
            Width           =   675
         End
      End
   End
   Begin MSComctlLib.TabStrip sTab 
      Height          =   4035
      Left            =   135
      TabIndex        =   7
      Top             =   150
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   7117
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Form Designer"
            Key             =   "T_FORM"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Editor"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CboTmp1 As String, CboTmp2 As String

Public Sub LoadConfig()
Dim FormSelection As String, EditorSelection As String, FoundPos As Long

    DevXML.XMLLoadFormFile App_ConfigFile
    FormSelection = DevXML.GetSelection("FormDesigner")
    EditorSelection = DevXML.GetSelection("editor")
    
    If DevXML.HasError > 0 Then
        SaveDefaultSettings
        WriteXMLIni
    End If
    
    AppConfig.ShowGrid = CBool(DevXML.GetSelectionValue(FormSelection, "ShowGrid", "value", "TRUE"))
    AppConfig.GridX = CInt(DevXML.GetSelectionValue(FormSelection, "GridX", "value", "120"))
    AppConfig.GridY = CInt(DevXML.GetSelectionValue(FormSelection, "GridY", "value", "120"))
    AppConfig.GridColor = CLng(DevXML.GetSelectionValue(FormSelection, "GridColor", "value", "4210752"))
    AppConfig.MovableAtRunTime = CBool(DevXML.GetSelectionValue(FormSelection, "MovableAtRunTime", "value", "FALSE"))
    AppConfig.ShowControlHints = CBool(DevXML.GetSelectionValue(FormSelection, "ShowControlHints", "value", "TRUE"))
    
    AppConfig.EditFont = DevXML.GetSelectionValue(EditorSelection, "Font", "value", "Courier New")
    AppConfig.EditFontSize = CInt(DevXML.GetSelectionValue(EditorSelection, "FontSize", "value", "10"))
    AppConfig.EditFontColor = CLng(DevXML.GetSelectionValue(EditorSelection, "FontColour", "value", "0"))
    AppConfig.EditFontBackColor = CLng(DevXML.GetSelectionValue(EditorSelection, "EditFontBackColor", "value", "16777215"))
    AppConfig.EditTabSize = CInt(DevXML.GetSelectionValue(EditorSelection, "TabSize", "value", "4"))
    AppConfig.EditCodeinsight = CBool(DevXML.GetSelectionValue(EditorSelection, "Codeinsight", "value", "TRUE"))
    AppConfig.EditMarginBar = CBool(DevXML.GetSelectionValue(EditorSelection, "EditMarginBar", "value", "TRUE"))

    chkGrid.Value = Abs(AppConfig.ShowGrid)
    txtWidth.Text = AppConfig.GridX
    txtHeight.Text = AppConfig.GridY
    picGridCol.BackColor = AppConfig.GridColor
    chkFormMove.Value = Abs(AppConfig.MovableAtRunTime)
    chkhints.Value = Abs(AppConfig.ShowControlHints)
    
    For I = 0 To cboFont.ListCount
        If LCase(AppConfig.EditFont) = LCase(cboFont.List(I)) Then
            FoundPos = I
            Exit For
        End If
    Next
    cboFont.ListIndex = FoundPos: FoundPos = 0
    For I = 0 To cbofontsize.ListCount
        If LCase(AppConfig.EditFontSize) = LCase(cbofontsize.List(I)) Then
            FoundPos = I
            Exit For
        End If
    Next
    I = 0: cbofontsize.ListIndex = FoundPos: FoundPos = 0
    
    PicEdFCol.BackColor = AppConfig.EditFontColor
    PicEdBkCol.BackColor = AppConfig.EditFontBackColor
    txtTabSize.Text = AppConfig.EditTabSize
    chkcode.Value = Abs(AppConfig.EditCodeinsight)
    chkMargin.Value = Abs(AppConfig.EditMarginBar)
    PicMargin.Visible = chkMargin
    
    If DevXML.HasError > 0 Then
        SaveDefaultSettings
        WriteXMLIni
    End If
    
    EditorSelection = ""
    FormSelection = ""
    DevXML.XMLReset
End Sub

Private Sub ArrangleTabs(Index As Integer)
Dim I As Integer
    For I = 0 To PicTab1.Count - 1
        PicTab1(I).Visible = False
    Next
    I = 0
    
    PicTab1(Index).Top = 540
    PicTab1(Index).Left = 255
    PicTab1(Index).Visible = True
    
End Sub

Sub DoColor(PicBox As PictureBox)

On Error GoTo CanErr
    With CDLG
        .CancelError = True
        .ShowColor
        PicBox.BackColor = .Color
    End With
CanErr:
    If Err = cdlCancel Then Exit Sub
    
End Sub
Private Sub cboFont_Change()
    cboFont.Text = CboTmp1
End Sub

Private Sub cboFont_Click()
    CboTmp1 = cboFont.Text
End Sub

Private Sub cbofontsize_Change()
    cbofontsize.Text = CboTmp2
End Sub

Private Sub cbofontsize_Click()
    CboTmp2 = cbofontsize.Text
End Sub

Private Sub chkMargin_Click()
    PicMargin.Visible = chkMargin
    
End Sub

Private Sub cmdCancel_Click()
    cboFont.Clear
    cbofontsize.Clear
    CboTmp1 = ""
    CboTmp2 = ""
    Unload frmOptions
End Sub

Private Sub cmdok_Click()
   If IsNumeric(txtWidth.Text) = False Or IsNumeric(txtHeight.Text) = False Then
        MsgBox "Grid size must be in the value of 50 to 1180", vbCritical, "Inavild Grid Value"
        Exit Sub
    ElseIf Val(txtWidth.Text) < 50 Or Val(txtHeight.Text) < 50 Then
        MsgBox "Grid size must be in the value of 50 to 1180", vbCritical, "Inavild Grid Value"
        Exit Sub
    ElseIf IsNumeric(txtTabSize.Text) = False Then
        MsgBox "Tab size must be in the value of 1 to 32", vbCritical, "Inavild Value"
        Exit Sub
    ElseIf Val(txtTabSize.Text) <= 0 Or Val(txtTabSize.Text) > 32 Then
        MsgBox "Tab size must be in the value of 1 to 32", vbCritical, "Inavild Value"
    Else
        ' Form designer
        AppConfig.ShowGrid = chkGrid
        AppConfig.GridX = Val(txtWidth.Text)
        AppConfig.GridY = Val(txtHeight.Text)
        AppConfig.GridColor = picGridCol.BackColor
        AppConfig.MovableAtRunTime = chkFormMove
        AppConfig.ShowControlHints = chkhints
        ' Editor
        AppConfig.EditFont = CboTmp1
        AppConfig.EditFontSize = Val(CboTmp2)
        AppConfig.EditFontColor = PicEdFCol.BackColor
        AppConfig.EditFontBackColor = PicEdBkCol.BackColor
        AppConfig.EditTabSize = Val(txtTabSize.Text)
        AppConfig.EditCodeinsight = chkcode
        AppConfig.EditMarginBar = chkMargin
        WriteXMLIni
        frmCode.Form_Activate
        MDIForm1.nGrid = chkGrid
        MDIForm1.mnugrid_Click
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
Dim I As Integer
    Set frmOptions.Icon = Nothing

    For I = 0 To Screen.FontCount - 1
        cboFont.AddItem Screen.Fonts(I)
        If LCase(Screen.Fonts(I)) = "courier" Then cboFont.ListIndex = I
    Next
    I = 0
    cbofontsize.AddItem "8"
    cbofontsize.AddItem "9"
    cbofontsize.AddItem "10"
    cbofontsize.AddItem "11"
    cbofontsize.AddItem "12"
    cbofontsize.AddItem "14"
    cbofontsize.AddItem "16"
    cbofontsize.AddItem "18"
    cbofontsize.AddItem "24"
    cbofontsize.ListIndex = 2
    FlatBorder p1.hwnd
    sTab_Click
    LoadConfig
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOptions = Nothing
End Sub

Private Sub PicEdBkCol_Click()
    DoColor PicEdBkCol
End Sub

Private Sub PicEdFCol_Click()
    DoColor PicEdFCol
End Sub

Private Sub picGridCol_Click()
    DoColor picGridCol
End Sub

Private Sub sTab_Click()
    ArrangleTabs (sTab.SelectedItem.Index - 1)
End Sub
