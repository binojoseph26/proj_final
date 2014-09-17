VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00C0FFFF&
   Caption         =   "SARASWATI VIDYALAYA"
   ClientHeight    =   9705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16515
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   0
      Picture         =   "Form1.frx":7856D
      ScaleHeight     =   1875
      ScaleWidth      =   1845
      TabIndex        =   6
      Top             =   -120
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      ForeColor       =   &H00400000&
      Height          =   7215
      Left            =   0
      Picture         =   "Form1.frx":79631
      ScaleHeight     =   7155
      ScaleWidth      =   12315
      TabIndex        =   4
      Top             =   2520
      Width           =   12375
   End
   Begin VB.TextBox password 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   16800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox username 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16800
      TabIndex        =   0
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label manual 
      BackColor       =   &H00400000&
      Caption         =   "USER MANUAL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   495
      Left            =   17280
      TabIndex        =   14
      Top             =   9360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      X1              =   2640
      X2              =   2640
      Y1              =   2520
      Y2              =   10920
   End
   Begin VB.Label Label9 
      BackColor       =   &H00400000&
      Caption         =   "Label8"
      Height          =   8415
      Left            =   2400
      TabIndex        =   13
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Height          =   8415
      Left            =   2040
      TabIndex        =   12
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Height          =   8415
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Product Info"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   17280
      TabIndex        =   10
      Top             =   9960
      Width           =   2655
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
      TabIndex        =   9
      Top             =   960
      Width           =   4575
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
      TabIndex        =   8
      Top             =   0
      Width           =   18375
   End
   Begin VB.Label login_label 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   17160
      TabIndex        =   7
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   20535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   " Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   14400
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   " User Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   14400
      TabIndex        =   2
      Top             =   3480
      Width           =   1815
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public loginflag As Integer


Private Sub username_load()
username.Text = ""
password.Text = ""
End Sub

Private Sub Command2_Click()
frmSplash.Show
login.Hide
End Sub


Private Sub Label6_Click()
Product_info.Show
login.Hide
Unload Me
End Sub

Private Sub login_label_Click()
If username.Text = "admin" And password.Text = "pict123" Then
loginflag = 1
login.Hide
menu.Show
Unload Me

ElseIf username.Text = "database" And password = "saraswati" Then
loginflag = 0
menu.Show
login.Hide
Unload Me
Else
loginflag = 1
abc = MsgBox("Wrong username or password!", vbOK, "Error")
End If
username.Text = ""
password.Text = ""
End Sub

Private Sub password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    If username.Text = "admin" And password.Text = "pict123" Then
loginflag = 1
login.Hide
menu.Show
Unload Me

ElseIf username.Text = "database" And password = "saraswati" Then
loginflag = 0
menu.Show
login.Hide
Unload Me
Else
loginflag = 1
abc = MsgBox("Wrong username or password!", vbOK, "Error")
End If
username.Text = ""
password.Text = ""
End If

End Sub

