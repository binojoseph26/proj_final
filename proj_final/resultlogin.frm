VERSION 5.00
Begin VB.Form resultlogin 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3165
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "resultlogin.frx":0000
   ScaleHeight     =   1869.986
   ScaleMode       =   0  'User
   ScaleWidth      =   6309.716
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1290
      TabIndex        =   1
      Top             =   120
      Width           =   3645
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  &PRN:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "resultlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public login As String
Public LoginSucceeded As Boolean


Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    RESULT.Show
End Sub

Private Sub cmdOK_Click()
login = txtUserName.Text
RESULT.Label2.Visible = True
RESULT.examtype.Visible = True
RESULT.Show
Me.Hide
End Sub


