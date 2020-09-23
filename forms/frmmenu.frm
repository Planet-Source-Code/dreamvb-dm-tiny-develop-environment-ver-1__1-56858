VERSION 5.00
Begin VB.Form frmmenu 
   ClientHeight    =   75
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   75
   ScaleWidth      =   1560
   Begin VB.Menu Mnu1 
      Caption         =   "Top"
      Begin VB.Menu mnudeleteA 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnufront 
         Caption         =   "&Bring to Front"
      End
      Begin VB.Menu mnuback 
         Caption         =   "&Send to Back"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprop 
         Caption         =   "&Properties"
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuback_Click()
    TheObjectName.ZOrder vbSendToBack
End Sub

Private Sub mnudeleteA_Click()
    frmWorkArea.HideSelection
    Unload TheObjectName
End Sub

Private Sub mnufront_Click()
    TheObjectName.ZOrder vbBringToFront
End Sub
