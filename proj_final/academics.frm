VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form academics 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   Caption         =   "Academics"
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   16110
   LinkTopic       =   "Form2"
   Picture         =   "academics.frx":0000
   ScaleHeight     =   7710
   ScaleWidth      =   16110
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   0
      Picture         =   "academics.frx":7856D
      ScaleHeight     =   1875
      ScaleWidth      =   1845
      TabIndex        =   23
      Top             =   0
      Width           =   1905
   End
   Begin VB.CommandButton Home1 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3120
      Picture         =   "academics.frx":79631
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton backbutton 
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
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
      Index           =   5
      Left            =   9960
      TabIndex        =   18
      Top             =   9600
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   4800
      TabIndex        =   17
      Top             =   5160
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   14737632
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   33
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Searchbutton 
      BackColor       =   &H000040C0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
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
      Index           =   4
      Left            =   9960
      TabIndex        =   15
      Top             =   8760
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
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
      Index           =   3
      Left            =   9960
      TabIndex        =   14
      Top             =   7920
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
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
      Index           =   2
      Left            =   9960
      TabIndex        =   13
      Top             =   7080
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
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
      Index           =   1
      Left            =   9960
      TabIndex        =   12
      Top             =   6240
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   495
      Index           =   0
      Left            =   9960
      TabIndex        =   6
      Top             =   5400
      Width           =   2895
   End
   Begin VB.ComboBox category 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9720
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3720
      Width           =   3135
   End
   Begin VB.OptionButton check 
      BackColor       =   &H000040C0&
      Height          =   255
      Index           =   1
      Left            =   9120
      TabIndex        =   2
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton check 
      BackColor       =   &H000040C0&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   1
      Top             =   3000
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      X1              =   2520
      X2              =   2520
      Y1              =   2640
      Y2              =   10920
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Height          =   8415
      Left            =   2280
      TabIndex        =   29
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   1920
      TabIndex        =   28
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   1440
      TabIndex        =   27
      Top             =   2640
      Width           =   375
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   0
      Width           =   18375
   End
   Begin VB.Label title 
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   24
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Caption         =   "Label3"
      Height          =   735
      Left            =   -120
      TabIndex        =   22
      Top             =   1920
      Width           =   20775
   End
   Begin VB.Label title 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Index           =   6
      Left            =   8160
      TabIndex        =   19
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label title 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Index           =   5
      Left            =   8160
      TabIndex        =   11
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label title 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
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
      Index           =   4
      Left            =   8160
      TabIndex        =   10
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label title 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   8160
      TabIndex        =   9
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label title 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Roll No"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   8160
      TabIndex        =   8
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label title 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRN"
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
      Index           =   1
      Left            =   8160
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label search 
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Index           =   2
      Left            =   7440
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label search 
      BackColor       =   &H000040C0&
      Caption         =   "Search by Roll number"
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
      Index           =   1
      Left            =   9720
      TabIndex        =   3
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label search 
      BackColor       =   &H000040C0&
      Caption         =   "Search by PRN"
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
      Left            =   5760
      TabIndex        =   0
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Menu stud_info 
      Caption         =   "Student info"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu department 
      Caption         =   "Department"
      Begin VB.Menu secondary 
         Caption         =   "Secondary"
         Begin VB.Menu five 
            Caption         =   "class 5"
         End
         Begin VB.Menu six 
            Caption         =   "class 6"
         End
         Begin VB.Menu seven 
            Caption         =   "class 7"
         End
         Begin VB.Menu eight 
            Caption         =   "class 8"
         End
         Begin VB.Menu nine 
            Caption         =   "class 9"
         End
         Begin VB.Menu ten 
            Caption         =   "class 10"
         End
      End
      Begin VB.Menu high_sec 
         Caption         =   "Higher secondary"
         Begin VB.Menu eleven 
            Caption         =   "class 11"
         End
         Begin VB.Menu twelve 
            Caption         =   "class 12"
         End
      End
   End
   Begin VB.Menu Fac_info 
      Caption         =   "Faculty info"
      Begin VB.Menu Fac_sec 
         Caption         =   "Secondary"
      End
      Begin VB.Menu fac_high_sec 
         Caption         =   "Higher secondary"
      End
   End
   Begin VB.Menu datareport 
      Caption         =   "Datareport"
   End
End
Attribute VB_Name = "academics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim recset As New ADODB.Recordset
Dim sroll, sprn, okay, temp As String
Dim checkflag As Double



Private Sub backbutton_Click()
okay = clear()
End Sub

Private Sub category_click()
title(0).Caption = category.Text
End Sub

Private Sub check_Click(Index As Integer)
category.Visible = True
Searchbutton(0).Visible = True
If Index = 0 Then
sprn = InputBox("Enter the PRN to be searched..", "Search by PRN", "")
checkflag = 0
ElseIf Index = 1 Then
sroll = InputBox("Enter the Roll no. to be searched..", "Search by Roll no", "")
checkflag = 1
End If
If sprn = "" Then
check(Index).Value = False
End If
End Sub

Private Sub datareport_Click()
DataReport1.Show
End Sub

Private Sub Home1_Click()
academics.Hide
menu.Show
connect.close
Unload Me
End Sub

Private Sub eight_Click()
sql = "select spers.name,sacad.prn,sacad.class,sacad.div,sacad.marks,sacad.att from sacademic sacad inner join spersonal spers on spers.prn=sacad.prn where sacad.class='8th'"
okay = execute_grid(sql)
End Sub

Private Sub eleven_Click()
sql = "select spers.name,sacad.prn,sacad.class,sacad.div,sacad.marks,sacad.att from sacademic sacad inner join spersonal spers on spers.prn=sacad.prn where sacad.class='11th'"
okay = execute_grid(sql)
End Sub

Private Sub fac_high_sec_Click()
sql = "select t.id,t.name,t.email,p.qual,p.leaves from tpersonal t,toffice p where p.school='hsecondary' And t.id = p.id"
okay = execute_grid(sql)
End Sub

Private Sub Fac_sec_Click()
sql = "select t.id,t.name,t.email,p.qual,p.leaves from tpersonal t,toffice p where p.school='secondary' And t.id = p.id"
okay = execute_grid(sql)
End Sub

Private Sub five_Click()
sql = "select spers.name,sacad.prn,sacad.class,sacad.div,sacad.marks,sacad.att from sacademic sacad inner join spersonal spers on spers.prn=sacad.prn where sacad.class='5th'"
okay = execute_grid(sql)
End Sub

Private Sub Form_Load()
Set connect = New ADODB.Connection
connect.Open "newmethod", "system", "pict1228"
sroll = ""
sprn = ""
okay = clear()
temp = ""
End Sub

Private Sub nine_Click()
sql = "select spers.name,sacad.prn,sacad.class,sacad.div,sacad.marks,sacad.att from sacademic sacad inner join spersonal spers on spers.prn=sacad.prn where sacad.class='9th'"
okay = execute_grid(sql)
End Sub


Private Sub Searchbutton_Click(Index As Integer)
backbutton.Visible = True
If check(1).Value = True And sroll = "" Then
okay = MsgBox("Please enter Roll no!", vbOKOnly)
Exit Sub
End If


If check(0).Value = True And sprn = "" Then
okay = MsgBox("Please enter PRN!", vbOKOnly)
Exit Sub
End If

If category.Text = "" Then
okay = MsgBox("Please select a category!", vbOKOnly)
Exit Sub
End If

If category.Text = "Personal" Then
title(0).Caption = "PERSONAL"
title(1).Caption = "PRN"
title(2).Caption = "Roll no."
title(3).Caption = "Name"
title(4).Caption = "DOB"
title(5).Caption = "Address"
title(6).Caption = "Phone"

Set recset = New ADODB.Recordset

If checkflag = 0 Then   'search using prn
sql = "select roll from sacademic where prn='" + sprn + "'"
Set recset = connect.Execute(sql)
           If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
sroll = recset(0)
sql = "select * from spersonal where prn='" + sprn + "'"

ElseIf checkflag = 1 Then   'search using roll no
sql = "select prn from sacademic where roll='" + sroll + "'"
Set recset = connect.Execute(sql)
If recset.EOF Then
MsgBox "record not found!", vbOKOnly
Exit Sub
End If
Set recset = connect.Execute(sql)
sql = "select * from spersonal where prn ='" + recset(0) + "'"
End If

recset.close
Set recset = connect.Execute(sql)
connect.Execute "commit"
For i = 0 To 6
title(i).Visible = True
Next
For i = 0 To 5
Text1(i).Visible = True
Next
 
Text1(0).Text = recset(4)
Text1(1).Text = sroll
Text1(2).Text = recset(0)
Text1(3).Text = recset(1)
Text1(4).Text = recset(2)
Text1(5).Text = recset(3)
MsgBox "Searched Successfully", vbOKOnly
recset.close

ElseIf category.Text = "Academic" Then
For i = 0 To 6
title(i).Visible = True
Next
For i = 0 To 5
Text1(i).Visible = True
Next

title(0).Caption = "ACADEMICS"
title(1).Caption = "PRN"
title(2).Caption = "Roll no."
title(3).Caption = "Class"
title(4).Caption = "Div"
title(5).Caption = "Marks"
title(6).Caption = "Attendance"
For i = 0 To 5
Text1(i).Visible = True
Next
Set recset = New ADODB.Recordset
If checkflag = 0 Then
sql = "select * from sacademic where prn='" + sprn + "'"
Else
sql = "select * from sacademic where roll='" + sroll + "'"
End If

Set recset = connect.Execute(sql)
connect.Execute "commit"
           If recset.EOF Then
           For i = 0 To 6
            title(i).Visible = False
            Next
            For i = 0 To 5
            Text1(i).Visible = False
            Next
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
For i = 0 To 5
Text1(i).Text = recset(i)
Next
MsgBox "Searched Successfully", vbOKOnly
recset.close

End If


End Sub

Private Sub seven_Click()
sql = "select spers.name,sacad.prn,sacad.class,sacad.div,sacad.marks,sacad.att from sacademic sacad inner join spersonal spers on spers.prn=sacad.prn where sacad.class='7th'"
okay = execute_grid(sql)
End Sub

Private Sub six_Click()
sql = "select spers.name,sacad.prn,sacad.class,sacad.div,sacad.marks,sacad.att from sacademic sacad inner join spersonal spers on spers.prn=sacad.prn where sacad.class='6th'"
okay = execute_grid(sql)
End Sub

Private Sub stud_info_Click()
search(0).Visible = True
search(1).Visible = True
search(2).Visible = False
check(0).Visible = True
check(1).Visible = True
category.List(0) = "Personal"
category.List(1) = "Academic"
DataGrid1.Visible = False
End Sub

Private Sub ten_Click()
sql = "select spers.name,sacad.prn,sacad.class,sacad.div,sacad.marks,sacad.att from sacademic sacad inner join spersonal spers on spers.prn=sacad.prn where sacad.class='10th'"
okay = execute_grid(sql)
End Sub


Private Sub twelve_Click()
sql = "select spers.name,sacad.prn,sacad.class,sacad.div,sacad.marks,sacad.att from sacademic sacad inner join spersonal spers on spers.prn=sacad.prn where sacad.class='12th'"
okay = execute_grid(sql)
End Sub

Function clear()
search(0).Visible = False
search(1).Visible = False
search(2).Visible = False
category.Visible = False
check(0).Visible = False
check(1).Visible = False
check(0).Value = False
check(1).Value = False

For i = 0 To 6
title(i).Visible = False
Next
For i = 0 To 5
Text1(i).Visible = False
Next
Searchbutton(0).Visible = False
DataGrid1.Visible = False
backbutton.Visible = False
End Function

Function execute_grid(sql)
okay = clear()
backbutton.Visible = True
Set recset = New ADODB.Recordset
recset.CursorType = adOpenStatic
recset.CursorLocation = adUseClient
recset.LockType = adLockOptimistic
recset.Open sql, connect
           If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Function
            End If
DataGrid1.Visible = True
Set DataGrid1.DataSource = recset
DataGrid1.Refresh
DataGrid1.Font.Bold = True
DataGrid1.BorderStyle = dbgNoBorder

End Function
