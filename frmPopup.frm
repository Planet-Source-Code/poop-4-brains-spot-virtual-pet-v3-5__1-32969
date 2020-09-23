VERSION 5.00
Begin VB.Form frmPopup 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   4680
   Icon            =   "frmPopup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu POP_UP 
      Caption         =   "Pop"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Hide Window"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
AddToTray frmSpot, "Spot v3.0", Me.Icon
End Sub

Private Sub mnuAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
Call frmSpot.cmdExit_Click
Unload frmAbout
Unload Me
Unload frmHelp
End Sub

Private Sub mnuOpen_Click()
Select Case mnuOpen.Caption
Case "Open Window"
frmSpot.Visible = True
mnuOpen.Caption = "Hide Window"
Case "Hide Window"
frmSpot.Visible = False
mnuOpen.Caption = "Open Window"
End Select
End Sub
