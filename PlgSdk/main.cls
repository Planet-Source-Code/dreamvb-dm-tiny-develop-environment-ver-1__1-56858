VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' You can use this project as a base to build your plug-ins

' mObject.DevCodeWindow will give you access the code editor
' mObject.DevIDE will give you access to the IDE it's self

Public Sub RunPlugin(mObject As Object)
    Set Form1.Editor = mObject.DevCodeWindow.txtCode
    Set Form1.DevIDE = mObject.DevIDE '
    Form1.Show
End Sub
