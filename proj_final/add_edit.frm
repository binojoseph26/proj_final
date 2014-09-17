VERSION 5.00
Begin VB.Form add_edit 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12285
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "add_edit.frx":0000
   ScaleHeight     =   9510
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   -120
      Picture         =   "add_edit.frx":7856D
      ScaleHeight     =   1875
      ScaleWidth      =   1845
      TabIndex        =   36
      Top             =   0
      Width           =   1905
   End
   Begin VB.CommandButton Home1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Picture         =   "add_edit.frx":79631
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   13680
      TabIndex        =   34
      Top             =   9600
      Width           =   3135
   End
   Begin VB.CommandButton entity 
      BackColor       =   &H000040C0&
      Caption         =   "Activities"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   16080
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   13680
      TabIndex        =   31
      Top             =   8760
      Width           =   3135
   End
   Begin VB.CommandButton save 
      BackColor       =   &H000040C0&
      Caption         =   "Submit"
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
      Index           =   1
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   13680
      TabIndex        =   28
      Top             =   7920
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   13680
      TabIndex        =   27
      Top             =   7080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   13680
      TabIndex        =   26
      Top             =   6240
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   13680
      TabIndex        =   25
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   7200
      TabIndex        =   24
      Top             =   9600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   7200
      TabIndex        =   23
      Top             =   8760
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   7200
      TabIndex        =   22
      Top             =   7920
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   7200
      TabIndex        =   21
      Top             =   7080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7200
      TabIndex        =   20
      Top             =   6240
      Width           =   3135
   End
   Begin VB.CommandButton entity 
      BackColor       =   &H000040C0&
      Caption         =   "Student"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton entity 
      BackColor       =   &H000040C0&
      Caption         =   "Teacher"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton save 
      BackColor       =   &H000040C0&
      Caption         =   "Save"
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
      Index           =   0
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10440
      Width           =   1695
   End
   Begin VB.CommandButton options 
      BackColor       =   &H000040C0&
      Caption         =   "Back"
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
      Index           =   3
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   7200
      TabIndex        =   5
      Top             =   5400
      Width           =   3135
   End
   Begin VB.CommandButton options 
      BackColor       =   &H000040C0&
      Caption         =   "Delete"
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
      Index           =   2
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton options 
      BackColor       =   &H000040C0&
      Caption         =   "Update"
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
      Index           =   1
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton options 
      BackColor       =   &H000040C0&
      Caption         =   "Add"
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
      Index           =   0
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton entity 
      BackColor       =   &H000040C0&
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton entity 
      BackColor       =   &H000040C0&
      Caption         =   "Accounts"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Height          =   735
      Left            =   -240
      TabIndex        =   42
      Top             =   1920
      Width           =   20775
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
      Left            =   2040
      TabIndex        =   41
      Top             =   0
      Width           =   18375
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
      TabIndex        =   40
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   1320
      TabIndex        =   39
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   1800
      TabIndex        =   38
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Height          =   8415
      Left            =   2160
      TabIndex        =   37
      Top             =   2640
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      X1              =   2400
      X2              =   2400
      Y1              =   2640
      Y2              =   10920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Paid (Term-2)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   11280
      TabIndex        =   33
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Paid (Term-1)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   11280
      TabIndex        =   30
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   11280
      TabIndex        =   19
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   11280
      TabIndex        =   18
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   11280
      TabIndex        =   17
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   11280
      TabIndex        =   16
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4800
      TabIndex        =   15
      Top             =   9600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Marks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4800
      TabIndex        =   14
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Div"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   4800
      TabIndex        =   13
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Roll"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   11
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRN"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   6
      Top             =   5400
      Width           =   1935
   End
End
Attribute VB_Name = "add_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recset As ADODB.Recordset
Dim connect As ADODB.Connection
Dim entityflag, blankfield As Double
Dim sprn, n1, tclass, csql As String
Dim tid, oid, oid1 As Double
Dim dob As String


Private Sub entity_Click(Index As Integer)
options(0).Visible = True
options(1).Visible = True
options(2).Visible = True
options(0).Enabled = True
options(1).Enabled = True
options(2).Enabled = True
options(3).Visible = True

For i = 0 To 4
entity(i).Enabled = False
Next

If Index = 0 Then
    entityflag = 0
ElseIf Index = 1 Then
    entityflag = 1
ElseIf Index = 2 Then
options(2).Visible = False
    entityflag = 2
ElseIf Index = 3 Then
    entityflag = 3
ElseIf Index = 4 Then
    entityflag = 4
    
End If
End Sub

Private Sub Form_Load()
Set connect = New ADODB.Connection
connect.Open "newmethod", "system", "pict1228"
For i = 0 To 11
Label1(i).BackStyle = 0
Label1(i).Visible = False
Next
For i = 0 To 11
Text1(i).Visible = False
Text1(i).BackColor = &HFFFFFF
Next
For i = 0 To 3
options(i).Visible = False
Next
save(0).Visible = False
save(1).Visible = False
End Sub



Private Sub Home1_Click()
menu.Show
add_edit.Hide
Unload Me

End Sub

Private Sub Label2_Click()

End Sub

Private Sub options_Click(Index As Integer)
If Index = 0 Then   'for add
    Text1(10).Visible = False
    Label1(10).Visible = False
    options(1).Visible = False
    options(2).Visible = False
    options(0).Visible = False
    options(3).Visible = True

If entityflag = 0 Then
    For i = 0 To 11
    Label1(i).Visible = True
    Next
    For i = 0 To 11
    Text1(i).Visible = True
    Next
    save(0).Visible = True
    Label1(0).Caption = "PRN"
    Label1(1).Caption = "Roll"
    Label1(2).Caption = "class"
    Label1(3).Caption = "Div"
    Label1(4).Caption = "Marks"
    Label1(5).Caption = "Attendance"
    Label1(6).Caption = "Name"
    Label1(7).Caption = "DOB"
    Label1(8).Caption = "Address"
    Label1(9).Caption = "Phone"
    Label1(10).Caption = "Fees paid (Term-1)"
    Label1(11).Caption = "Fees paid (term-2)"


ElseIf entityflag = 1 Then
    save(0).Visible = True
    Label1(0).Caption = "ID"
    Label1(1).Caption = "Name"
    Label1(2).Caption = "DOB"
    Label1(3).Caption = "Address"
    Label1(4).Caption = "Phone"
    Label1(5).Caption = "Email"
    Label1(6).Caption = "School"
    Label1(7).Caption = "Join date"
    Label1(8).Caption = "Qual"
    Label1(9).Caption = "Salary"
    Label1(10).Caption = "Leaves"
    For i = 0 To 10
    Label1(i).Visible = True
    Next

    For i = 0 To 10
    Text1(i).Visible = True
    Next

ElseIf entityflag = 2 Then
    Label1(0).Caption = "Class"
    Label1(1).Caption = "Tution Fee"
    Label1(2).Caption = "Library Fee"
    Label1(3).Caption = "S Fee"
    Label1(4).Caption = "Bus fee"
    Label1(5).Caption = "Stationery Fee"
    For i = 0 To 5
    Label1(i).Visible = True
    Next
    For i = 0 To 5
    Text1(i).Visible = True
    Next
    save(0).Visible = True

ElseIf entityflag = 3 Then

    Label1(0).Caption = "PRN"
    Label1(1).Caption = "Test name"
    Label1(2).Caption = "Maths"
    Label1(3).Caption = "Science"
    Label1(4).Caption = "Social sciences"
    Label1(5).Caption = "English"
    Label1(6).Caption = "Hindi"
    Label1(7).Caption = "Marathi"
    save(0).Visible = True

    For i = 0 To 7
    Label1(i).Visible = True
    Next

    For i = 0 To 7
    Text1(i).Visible = True
    Next

ElseIf entityflag = 4 Then

    Label1(0).Caption = "Event Id"
    Label1(1).Caption = "Co-ordinator Id"
    Label1(2).Caption = "Event"
      Label1(3).Caption = "Proposed Date"
    Label1(4).Caption = "Participants"

    For i = 0 To 4
    Label1(i).Visible = True
    Next

    For i = 0 To 4
    Text1(i).Visible = True
    Next
    save(0).Visible = True
End If

ElseIf Index = 2 Then
    options(0).Enabled = False
    options(2).Enabled = False
    options(3).Visible = True

If entityflag = 0 Then
    
    sprn = InputBox("Enter the PRN of student..", "DELETE")
    If sprn <> "" Then
    Set recset = New ADODB.Recordset
    recset.CursorType = adOpenStatic
    recset.CursorLocation = adUseClient
    recset.LockType = adLockOptimistic
    
    sql = "delete from sacademic where prn='" + sprn + "'"
    okay = MsgBox("Delete?", vbOKCancel)
    If okay = 1 Then
        recset.Open sql, connect
        'Set recset = connect.Execute(sql)
        connect.Execute "commit"
            If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
        okay = MsgBox("Record " + sprn + " deleted!")
    End If
     'select CONSTRAINT_NAME from USER_CONSTRAINTs where TABLE_NAME='CLASS'
    End If
    
ElseIf entityflag = 1 Then
    tid = InputBox("Enter the ID of teacher..", "DELETE")
    If tid <> 0 Then
    Set recset = New ADODB.Recordset
    recset.CursorType = adOpenStatic
    recset.CursorLocation = adUseClient
    recset.LockType = adLockOptimistic
    
    sql = "delete from tpersonal where id=" & tid
    okay = MsgBox("Delete?", vbOKCancel)
    If okay = 1 Then
        recset.Open sql, connect
        connect.Execute "commit"
            If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
        okay = MsgBox("Record deleted!")
    End If
     'select CONSTRAINT_NAME from USER_CONSTRAINTs where TABLE_NAME='CLASS'
    End If
    
ElseIf entityflag = 3 Then
    sprn = InputBox("Enter the PRN of student..", "DELETE")
    If sprn <> "" Then
    Set recset = New ADODB.Recordset
    recset.CursorType = adOpenStatic
    recset.CursorLocation = adUseClient
    recset.LockType = adLockOptimistic
    
    sql = "delete from result where prn='" + sprn + "'"
    okay = MsgBox("Delete?", vbOKCancel)
    If okay = 1 Then
        recset.Open sql, connect
        connect.Execute "commit"
            If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
        okay = MsgBox("Record " + sprn + " deleted!")
    End If
     'select CONSTRAINT_NAME from USER_CONSTRAINTs where TABLE_NAME='CLASS'
    End If

ElseIf entityflag = 4 Then
    oid = InputBox("Enter the Event Id..", "DELETE")
    If oid <> 0 Then
    Set recset = New ADODB.Recordset
    recset.CursorType = adOpenStatic
    recset.CursorLocation = adUseClient
    recset.LockType = adLockOptimistic
    
    sql = "delete from other where eid= " & oid
    okay = MsgBox("Delete?", vbOKCancel)
    If okay = 1 Then
        recset.Open sql, connect
        connect.Execute "commit"
            If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
        okay = MsgBox("Event deleted!")
    End If
     'select CONSTRAINT_NAME from USER_CONSTRAINTs where TABLE_NAME='CLASS'
    End If

End If



ElseIf Index = 1 Then   'update option
    options(0).Visible = False
    options(2).Visible = False
    options(3).Visible = True
    
    If entityflag = 0 Then 'student update
    sprn = InputBox("Enter the PRN you want to update...!", "UPDATION")
    If sprn <> "" Then
        For i = 0 To 11
            Label1(i).Visible = True
        Next
        For i = 0 To 11
            Text1(i).Visible = True
        Next
        Text1(0).Enabled = False

        save(1).Visible = True
        Set recset = New ADODB.Recordset
        sql = "select * from sacademic where prn='" + sprn + "'"
        recset.CursorType = adOpenStatic
        recset.CursorLocation = adUseClient
        recset.LockType = adLockOptimistic
        recset.Open sql, connect
            If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
        For i = 0 To 5
            Text1(i).Text = recset(i)
        Next
        recset.close
        sql = "select * from spersonal where prn='" + sprn + "'"
        recset.Open sql, connect
        For i = 6 To 9
            Text1(i).Text = recset(i - 6)
        Next
        recset.close
        sql = "select * from feestatus where prn='" + sprn + "'"
        recset.Open sql, connect
        Text1(10).Text = recset(1)
        Text1(11).Text = recset(2)
        connect.Execute "commit"
        
        options(3).Visible = True
        recset.close
     End If
     
     ElseIf entityflag = 1 Then 'teacher update
        
        n1 = InputBox("Enter the ID you want to update...!", "UPDATION")
        tid = Val(n1)
        If tid <> 0 Then
        Label1(0).Caption = "ID"
        Label1(1).Caption = "Name"
        Label1(2).Caption = "DOB"
        Label1(3).Caption = "Address"
        Label1(4).Caption = "Phone"
        Label1(5).Caption = "Email"
        Label1(6).Caption = "School"
        Label1(7).Caption = "Join date"
        Label1(8).Caption = "Qual"
        Label1(9).Caption = "Salary"
        Label1(10).Caption = "Leaves"
        For i = 0 To 10
            Label1(i).Visible = True
        Next

        For i = 0 To 10
            Text1(i).Visible = True
        Next
        Text1(0).Enabled = False
        save(1).Visible = True
        Set recset = New ADODB.Recordset
        sql = "select * from tpersonal where id=" & tid
        recset.CursorType = adOpenStatic
        recset.CursorLocation = adUseClient
        recset.LockType = adLockOptimistic
        recset.Open sql, connect
            If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
        For i = 0 To 5
            Text1(i).Text = recset(i)
        Next
        recset.close
        sql = "select * from toffice where id=" & tid
        recset.Open sql, connect
        For i = 6 To 10
            Text1(i).Text = recset(i - 5)
        Next
        connect.Execute "commit"
        options(3).Visible = True
        recset.close
        End If
     
    ElseIf entityflag = 2 Then
    Label1(0).Caption = "Class"
    Label1(1).Caption = "Tution Fee"
    Label1(2).Caption = "Library Fee"
    Label1(3).Caption = "S Fee"
    Label1(4).Caption = "Bus fee"
    Label1(5).Caption = "Stationery Fee"
    For i = 0 To 5
    Label1(i).Visible = True
    Next
    For i = 0 To 5
    Text1(i).Visible = True
    Next
    tclass = InputBox("Enter the class..!!", "UPDATION")
    If tclass <> "" Then
        For i = 0 To 5
            Label1(i).Visible = True
        Next
        For i = 0 To 5
            Text1(i).Visible = True
        Next
        save(1).Visible = True
        Set recset = New ADODB.Recordset
        sql = "select * from feestr where class = '" & tclass & "'"
        recset.CursorType = adOpenStatic
        recset.CursorLocation = adUseClient
        recset.LockType = adLockOptimistic
        recset.Open sql, connect
            If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
        Text1(0).Text = recset(0)
        Text1(1).Text = recset(2)
        Text1(2).Text = recset(3)
        Text1(3).Text = recset(4)
        Text1(4).Text = recset(5)
        Text1(5).Text = recset(6)
        recset.close
     End If
     
     ElseIf entityflag = 3 Then 'result update
        
        sprn = InputBox("Enter the ID you want to update...!", "UPDATION")
        n1 = InputBox("Enter the testname!", "UPDATION")
        
        If sprn <> "" Then
        Label1(0).Caption = "PRN"
        Label1(1).Caption = "TEST Name"
        Label1(2).Caption = "MATHS"
        Label1(3).Caption = "SCIENCE"
        Label1(4).Caption = "SOCIAL SCIENCE"
        Label1(5).Caption = "ENGLISH"
        Label1(6).Caption = "HINDI"
        Label1(7).Caption = "MARATHI"
        For i = 0 To 7
            Label1(i).Visible = True
        Next

        For i = 0 To 7
            Text1(i).Visible = True
        Next
        Text1(0).Enabled = False
        save(1).Visible = True
        Set recset = New ADODB.Recordset
        sql = "select * from result where prn='" & sprn & "' and testname='" & n1 & "'"
        recset.CursorType = adOpenStatic
        recset.CursorLocation = adUseClient
        recset.LockType = adLockOptimistic
        recset.Open sql, connect
            If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
        Text1(0).Text = recset(7)
        For i = 0 To 6
            Text1(i + 1).Text = recset(i)
        Next
        options(3).Visible = True
        recset.close

        End If
        
      ElseIf entityflag = 4 Then 'event update
      Label1(0).Caption = "Event Id"
      Label1(1).Caption = "Co-ordinator Id"
      Label1(2).Caption = "Event"
      Label1(3).Caption = "Proposed Date"
      Label1(4).Caption = "Participants"

  
      oid = InputBox("Enter the Id of the event you want to update...!", "UPDATION")
       'oid1 = oid
       If oid <> 0 Then
             For i = 0 To 4
                Label1(i).Visible = True
            Next
            For i = 0 To 4
                Text1(i).Visible = True
            Next
            save(1).Visible = True
            Set recset = New ADODB.Recordset
            sql = "select * from other where eid =" & oid
            recset.CursorType = adOpenStatic
            recset.CursorLocation = adUseClient
            recset.LockType = adLockOptimistic
            recset.Open sql, connect
            If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
            For i = 0 To 4
                 Text1(i).Text = recset(i)
            Next
            recset.close
    End If

     
     
   End If
ElseIf Index = 3 Then 'Back
For i = 0 To 11
Label1(i).Visible = False
Next
For i = 0 To 11
Text1(i).Visible = False
Next
For i = 0 To 4
entity(i).Enabled = True
Next
For i = 0 To 3
options(i).Visible = False
Next
save(0).Visible = False
save(1).Visible = False
End If

End Sub

Private Sub save_Click(Index As Integer)

okay = MsgBox("Save the Data?", vbOKCancel)
If okay = vbOK Then


If Index = 0 Then   'index=0 for add
    If entityflag = 0 Then 'entityflag=0 for student
    For i = 0 To 11
    If Text1(i).Text = "" Then
    MsgBox "All fields are mandatory"
    Exit Sub
    End If
    Next
    sql = "insert into sacademic values('" & Text1(0).Text & "', '" & Text1(1).Text & "', '" & Text1(2).Text & "','" & Text1(3).Text & "'," & Text1(4).Text & ",'" & Text1(5).Text & "')"
    connect.Execute sql
    connect.Execute "commit"
    sql = "insert into spersonal values('" & Text1(6).Text & "','" & Text1(7).Text & "', '" & Text1(8).Text & "', " & Text1(9).Text & ",'" & Text1(0).Text & "')"
    connect.Execute sql
    connect.Execute "commit"
    sql = "insert into feestatus values('" & Text1(0).Text & "', " & Text1(10).Text & ", " & Text1(11).Text & ", '" & Null & "')"
    connect.Execute sql
    connect.Execute "commit"
     okay = MsgBox("Data inserted!", vbOKOnly)
    
    ElseIf entityflag = 1 Then 'entityflag=1 for teacher
    For i = 0 To 10
    If Text1(i).Text = "" Then
    MsgBox "All fields are mandatory"
    Exit Sub
    End If
    Next
    csql = "insert into tpersonal values(" & Text1(0).Text & ",'" & Text1(1).Text & "','" & Text1(2).Text & "','" & Text1(3).Text & "'," & Text1(4).Text & ",'" & Text1(5).Text & "')"
    connect.Execute csql
    connect.Execute "commit"

    csql = "insert into toffice values(" & Text1(0).Text & ", '" & Text1(6).Text & "', '" & Text1(7).Text & "','" & Text1(8).Text & "'," & Text1(9).Text & "," & Text1(10).Text & ")"
    connect.Execute csql
    connect.Execute "commit"
    okay = MsgBox("Data inserted!", vbOKOnly)
    
    ElseIf entityflag = 2 Then 'entityflag=2 for accounts
    total = 0
    For i = 0 To 5
    If Text1(i).Text = "" Then
    MsgBox "All fields are mandatory"
    Exit Sub
    End If
    Next
    For i = 1 To 4
    total = total + Val(Text1(i).Text)
    Next
    sql = "insert into feestr values('" & Text1(0).Text & "', 'A', " & Text1(1).Text & "," & Text1(2).Text & "," & Text1(3).Text & "," & Text1(4).Text & "," & Text1(5).Text & "," & total & ")"
    connect.Execute sql
    connect.Execute "commit"
    okay = MsgBox("Data inserted!", vbOKOnly)
     
    ElseIf entityflag = 3 Then
    For i = 0 To 7
    If Text1(i).Text = "" Then
    MsgBox "All fields are mandatory"
    Exit Sub
    End If
    Next
    sql = "insert into result values('" & Text1(1).Text & "', " & Text1(2).Text & "," & Text1(3).Text & "," & Text1(4).Text & "," & Text1(5).Text & "," & Text1(6).Text & "," & Text1(7).Text & ",'" & Text1(0).Text & "')"
    connect.Execute sql
    connect.Execute "commit"
    okay = MsgBox("Data inserted!", vbOKOnly)
    
    ElseIf entityflag = 4 Then
     For i = 0 To 4
    If Text1(i).Text = "" Then
    MsgBox "All fields are mandatory"
    Exit Sub
    End If
    Next
     sql = "insert into other values(" & Text1(0).Text & ", " & Text1(1).Text & ", '" & Text1(2).Text & "','" & Text1(3).Text & "'," & Text1(4).Text & ")"
    connect.Execute sql
    connect.Execute "commit"
    okay = MsgBox("Data inserted!", vbOKOnly)

     End If
    
    ElseIf Index = 1 Then   'index=1 for update
    If entityflag = 0 Then
    sql = "update sacademic set prn='" + Text1(0).Text + "',roll='" + Text1(1).Text + "',class='" + Text1(2).Text + "',div='" + Text1(3).Text + "',marks='" + Text1(4).Text + "',att='" + Text1(5).Text + " ' where prn='" + sprn + "'"
    connect.Execute sql
    connect.Execute "commit"

    sql = "update spersonal set name='" + Text1(6).Text + "',dob='" + Text1(7).Text + "',addr='" + Text1(8).Text + "',phone='" + Text1(9).Text + "' where prn='" + sprn + "'"
    connect.Execute sql
    connect.Execute "commit"
    sql = "update feestatus set term1=" + Text1(10).Text + ",term2=" + Text1(11).Text + " where prn='" + sprn + "'"
    connect.Execute sql
    connect.Execute "commit"
    okay = MsgBox("Data Updated!", vbOKOnly)
    
    ElseIf entityflag = 1 Then
    sql = "update tpersonal set id=" + Text1(0).Text + ",name='" + Text1(1).Text + "',dob='" + Text1(2).Text + "',address='" + Text1(3).Text + "',phone=" + Text1(4).Text + ",email='" + Text1(5).Text + " ' where id=" & tid
    connect.Execute sql
    connect.Execute "commit"
    sql = "update toffice set id=" + Text1(0).Text + ",school='" + Text1(6).Text + "',join_date='" + Text1(7).Text + "',qual='" + Text1(8).Text + "',salary=" + Text1(9).Text + ",leaves=" + Text1(10).Text + " where id=" & tid
    connect.Execute sql
    connect.Execute "commit"
    okay = MsgBox("Data Updated!", vbOKOnly)
    
     ElseIf entityflag = 2 Then
    'sql = "update feestr set class='" + Text1(0).Text + "', tfee=" + Text1(1).Text + ",lfee=" + Text1(2).Text + ",sfee=" + Text1(3).Text + ",bfee=" + Text1(4).Text + ", stfee=" + Text1(5).Text + " where class='" + tclass + "'"
    sql = "update feestr set tfee=" & Text1(1).Text & ",lfee=" & Text1(2).Text & ",sfee=" & Text1(3).Text & ",bfee=" & Text1(4).Text & ",stfee=" & Text1(5).Text & " where class='" & tclass & "' and div='A'"
    connect.Execute sql
    connect.Execute "commit"
    okay = MsgBox("Data Updated!", vbOKOnly)
    
    ElseIf entityflag = 3 Then
    sql = "update result set prn='" + Text1(0).Text + "',testname='" + Text1(1).Text + "',mat=" + Text1(2).Text + ",sci=" + Text1(3).Text + ",sst=" + Text1(4).Text + ",eng=" + Text1(5).Text + ",hin= " + Text1(6).Text + ",mar=" + Text1(7).Text + " where prn='" & sprn & "'"
    connect.Execute sql
    connect.Execute "commit"
    okay = MsgBox("Data Updated!", vbOKOnly)
    
    ElseIf entityflag = 4 Then
    sql = "update other set eid=" + Text1(0).Text + ",id=" + Text1(1).Text + ",event='" + Text1(2).Text + "',pdate='" + Text1(3).Text + "',participants=" + Text1(4).Text + "  where eid = " & oid
    connect.Execute sql
    connect.Execute "commit"
    okay = MsgBox("Data Updated!", vbOKOnly)
    End If
    
    For i = 0 To 10
    Text1(i).Text = ""
    Next
    Text1(0).Enabled = True
    End If
End If

End Sub

Private Sub Text1_Change(Index As Integer)
Dim temp As String
i = Index
If entityflag = 0 Then
If i = 4 Or i = 5 Or i = 9 Or i = 10 Or i = 11 Then
temp = Text1(i).Text
If (CStr(Val(temp)) <> temp) Then
MsgBox "Invalid Entry!"
Exit Sub
End If
End If



ElseIf entityflag = 1 Then
If i = 4 Or i = 9 Then
temp = Text1(i).Text
If (CStr(Val(temp)) <> temp) Then
MsgBox "Invalid Entry!"
Exit Sub
End If
End If


ElseIf entityflag = 2 Then
If i = 1 Or i = 2 Or i = 3 Or i = 4 Or i = 5 Then
temp = Text1(i).Text
If (CStr(Val(temp)) <> temp) Then
MsgBox "Invalid Entry!"
Exit Sub
End If
End If

ElseIf entityflag = 3 Then
If i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 6 Or i = 7 Then
temp = Text1(i).Text
If (CStr(Val(temp)) <> temp) Then
MsgBox "Invalid Entry!"
Exit Sub
End If
End If

End If
End Sub
