VERSION 5.00
Begin VB.Form frmMnu 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuAll 
      Caption         =   "All"
      Begin VB.Menu mnuSavePic 
         Caption         =   "Save Picture"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMini 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuabt_Click()

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuMini_Click()
frmMain.WindowState = vbMinimized
End Sub

Private Sub mnuSavePic_Click()
SavePicture frmMain.PBoxCat.Picture, "c:\CatOfDay.bmp"
End Sub
