VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public HasError As Boolean
Public DialogStrName As String
Public DialogObject As Form

Private TLanguage As String
Public sControl As ScriptControl
Private Graphics As clsGraphics
Private Windows As clswindows
Private FileSys As clsFiles
Private TNetwork As clsNetwork
Private TTools As clsTools
Private TIniFile As clsINI
Public ErrScript As Error

Public Sub SetupControl()

    Set sControl = New ScriptControl
    Set Graphics = New clsGraphics
    Set Windows = New clswindows
    Set FileSys = New clsFiles
    Set TNetwork = New clsNetwork
    Set TTools = New clsTools
    Set TIniFile = New clsINI
    
    sControl.Language = TLanguage
    sControl.AddObject DialogStrName, DialogObject, True
    sControl.AddObject "PicImg", frmWorkArea.PicImg, True
    sControl.AddObject "Graphics", Graphics, True
    sControl.AddObject "Windows", Windows, True
    sControl.AddObject "FileSys", FileSys, True
    sControl.AddObject "TNetwork", TNetwork, True
    sControl.AddObject "TTools", TTools, True
    sControl.AddObject "TIniFile", TIniFile, True
    sControl.AllowUI = True
    TTools.Hangle = DialogObject.hwnd
    Set ErrScript = sControl.Error
End Sub

Public Sub mRunProc(StrCode As String)
    On Error Resume Next
    sControl.Run StrCode
End Sub
Public Sub RunCode2(StrCode As String)
On Error GoTo CodeErr
    HasError = False
    sControl.AddCode StrCode
    Exit Sub
CodeErr:
    If Err Then
        If Err.Number = 13 Or Err.Number = 1004 Then Err.Clear
        HasError = True
    End If
End Sub
Public Sub Reset()
On Error Resume Next
    sControl.Reset
End Sub
Public Sub RunCode(StrCode As String)
On Error GoTo CodeErr
    HasError = False
    If Len(StrCode) = 0 Then Exit Sub
    sControl.AddCode StrCode
    Exit Sub
CodeErr:
    If Err Then
        If Err.Number = 13 Or Err.Number = 1004 Then Err.Clear
        HasError = True
        MDIForm1.GetLastScriptError
    End If
End Sub

Public Property Get mLanguage() As String
    mLanguage = TLanguage
End Property

Public Property Let mLanguage(ByVal vNewValue As String)
    TLanguage = vNewValue
End Property
