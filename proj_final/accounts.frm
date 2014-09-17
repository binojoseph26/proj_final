VERSION 5.00
Begin VB.Form accounts 
   Caption         =   "ACCOUNTS"
   ClientHeight    =   5955
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu home 
      Caption         =   "Home"
   End
   Begin VB.Menu fee_str 
      Caption         =   "Fee Structure"
   End
   Begin VB.Menu fee_sts 
      Caption         =   "Fee Status"
   End
End
Attribute VB_Name = "accounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fee_sts_Click()
acc_login.Show
accounts.Hide
End Sub

Private Sub home_Click()
menu.Show
faculty.Hide
End Sub
