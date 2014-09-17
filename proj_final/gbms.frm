VERSION 5.00
Begin VB.Form Product_info 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product info"
   ClientHeight    =   9570
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   19260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "gbms.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "gbms.frx":000C
   ScaleHeight     =   9570
   ScaleWidth      =   19260
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Height          =   6555
      Left            =   5760
      TabIndex        =   0
      Top             =   3480
      Width           =   11145
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "A product by"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pvt Ltd."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8400
         TabIndex        =   10
         Top             =   2640
         Width           =   975
      End
      Begin VB.Image imgLogo 
         Height          =   2625
         Left            =   480
         Picture         =   "gbms.frx":78579
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright @GBMS products"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   6240
         Width           =   3015
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRESS ENTER TO CONTINUE"
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
         Height          =   450
         Left            =   5520
         TabIndex        =   3
         Top             =   6000
         Width           =   4725
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Softwares"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   27.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   5520
         TabIndex        =   5
         Top             =   2280
         Width           =   2760
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Licensed to Sarswati Vidyalaya"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   3120
         TabIndex        =   1
         Top             =   4680
         Width           =   4215
      End
      Begin VB.Label lblBGMS 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "GBMS "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   36
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   3720
         TabIndex        =   4
         Top             =   1440
         Width           =   2385
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Height          =   735
      Index           =   9
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   20775
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   1680
      TabIndex        =   8
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   2160
      TabIndex        =   7
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Height          =   8415
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      X1              =   2760
      X2              =   2760
      Y1              =   2640
      Y2              =   10920
   End
End
Attribute VB_Name = "Product_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    login.Show
    Unload Me
End Sub

Private Sub lblCompany_Click()

End Sub
