VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form menu 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Welcome to Saraswati vidyalay!"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   Picture         =   "menu.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   0
      Picture         =   "menu.frx":697842
      ScaleHeight     =   1875
      ScaleWidth      =   1845
      TabIndex        =   11
      Top             =   0
      Width           =   1905
   End
   Begin VB.CommandButton close 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9240
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   864
      ImageHeight     =   588
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":698906
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":6AA4BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":6B3EE0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5280
      Top             =   4320
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   7815
      Left            =   3720
      ScaleHeight     =   7755
      ScaleWidth      =   12915
      TabIndex        =   7
      Top             =   3120
      Width           =   12975
   End
   Begin VB.CommandButton options 
      BackColor       =   &H00C0FFFF&
      Caption         =   "EDIT"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   7
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton options 
      BackColor       =   &H00C0FFFF&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   6
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton options 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ACCOUNTS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   5
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton options 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ACTIVITIES"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   3
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton options 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ACADEMICS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   2
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton options 
      BackColor       =   &H00C0FFFF&
      Caption         =   "INFO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   1080
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton options 
      BackColor       =   &H00C0FFFF&
      Caption         =   "RESULTS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   4
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Daulavadgoan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   30
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   855
      Left            =   9360
      TabIndex        =   15
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      X1              =   2640
      X2              =   2640
      Y1              =   2640
      Y2              =   10920
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Height          =   8415
      Left            =   2400
      TabIndex        =   14
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   2040
      TabIndex        =   13
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   1560
      TabIndex        =   12
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Saraswati Secondary and Higher Secondary School"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   975
      Left            =   2160
      TabIndex        =   10
      Top             =   0
      Width           =   18375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   20775
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub End1_Click()
End
End Sub

Private Sub close_Click()
End
End Sub

Private Sub Form_Load()
Picture1.Picture = ImageList1.ListImages(1).Picture
Timer1.Enabled = True
i = 2
If login.loginflag = 1 Then
options(7).Enabled = True
End If
End Sub
Private Sub Timer1_Timer()
Picture1.Picture = ImageList1.ListImages(i).Picture
i = ((i + 1) Mod 4)
If i = 0 Then
i = 1
End If
End Sub

Private Sub options_Click(Index As Integer)
Timer1.Enabled = False
If Index = 6 Then
login.Show
menu.Hide
ElseIf Index = 0 Then
Form2.Show
menu.Hide
ElseIf Index = 1 Then
faculty.Show
menu.Hide
ElseIf Index = 3 Then
Form1.Show
menu.Hide
ElseIf Index = 2 Then
academics.Show
menu.Hide
ElseIf Index = 5 Then
accounts.Show
menu.Hide
ElseIf Index = 4 Then
RESULT.Show
menu.Hide
ElseIf Index = 7 Then
add_edit.Show
menu.Hide

End If
Unload Me
End Sub


