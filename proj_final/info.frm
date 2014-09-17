VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8190
   LinkTopic       =   "Form2"
   Picture         =   "info.frx":0000
   ScaleHeight     =   5460
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   0
      Picture         =   "info.frx":697842
      ScaleHeight     =   1875
      ScaleWidth      =   1845
      TabIndex        =   3
      Top             =   0
      Width           =   1905
   End
   Begin VB.CommandButton Home1 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3600
      Picture         =   "info.frx":698906
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3720
      Width           =   8535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Height          =   8415
      Left            =   2400
      TabIndex        =   8
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   2040
      TabIndex        =   7
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   1560
      TabIndex        =   6
      Top             =   2640
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      X1              =   2640
      X2              =   2640
      Y1              =   2640
      Y2              =   10920
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
      Left            =   9480
      TabIndex        =   5
      Top             =   1080
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
      Left            =   2280
      TabIndex        =   4
      Top             =   0
      Width           =   18375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   20775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Text1.Text = "          SARASWATI VIDYALAYA The school founded in 1955 gives the students a divine place to build the foundation of their career with full devotion and sincerity and make them buil themselves for global competence and national character.The playful and gay environment of Saraswati vidyalaya helps the students to build and develop a competent personality.With best faculties in the state,the academic ambience helps the students come out with flying colours in all the competition the scholl takes part in."
End Sub

Private Sub Home1_Click()
Unload Me
menu.Show
End Sub

