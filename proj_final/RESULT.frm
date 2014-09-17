VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form result 
   BackColor       =   &H00C0FFFF&
   Caption         =   "RESULT"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17445
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "RESULT.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   17445
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   0
      Picture         =   "RESULT.frx":7856D
      ScaleHeight     =   1875
      ScaleWidth      =   1845
      TabIndex        =   37
      Top             =   0
      Width           =   1905
   End
   Begin VB.CommandButton Home1 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   4680
      Picture         =   "RESULT.frx":79631
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   6480
      TabIndex        =   34
      Top             =   6360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7646
      _Version        =   393216
      BackColor       =   12640511
      HeadLines       =   1
      RowHeight       =   28
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
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
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      Height          =   465
      Index           =   1
      Left            =   9240
      TabIndex        =   33
      Text            =   "Combo1"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      Height          =   465
      Index           =   0
      Left            =   4680
      TabIndex        =   32
      Text            =   "Combo1"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton back 
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
      Height          =   735
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton getresult 
      BackColor       =   &H000040C0&
      Caption         =   "Get result"
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5040
      Width           =   1695
   End
   Begin VB.ComboBox examtype 
      BackColor       =   &H000040C0&
      Height          =   465
      Left            =   7920
      TabIndex        =   28
      Text            =   "exam type"
      Top             =   5160
      Width           =   2655
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
      Height          =   615
      Index           =   12
      Left            =   13320
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   9960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   5
      EndProperty
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
      Index           =   11
      Left            =   9240
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   9960
      Width           =   1815
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
      Height          =   615
      Index           =   10
      Left            =   4680
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   9960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Index           =   9
      Left            =   13320
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   8880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Index           =   8
      Left            =   9240
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   8880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Index           =   7
      Left            =   4680
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   8760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Index           =   6
      Left            =   13320
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   7800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Index           =   5
      Left            =   9240
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   7800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Index           =   4
      Left            =   4680
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Index           =   3
      Left            =   17400
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Index           =   2
      Left            =   13200
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   9240
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   4680
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton overall 
      BackColor       =   &H000040C0&
      Caption         =   "Overall class result"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton sresult 
      BackColor       =   &H000040C0&
      Caption         =   "Student results"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00400000&
      X1              =   2640
      X2              =   2640
      Y1              =   2640
      Y2              =   10920
   End
   Begin VB.Label Label11 
      BackColor       =   &H00400000&
      Height          =   8415
      Left            =   2400
      TabIndex        =   42
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00400000&
      Height          =   8415
      Left            =   2040
      TabIndex        =   41
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00400000&
      Height          =   8295
      Left            =   1560
      TabIndex        =   40
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
      Left            =   9480
      TabIndex        =   39
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
      Left            =   2160
      TabIndex        =   38
      Top             =   0
      Width           =   18375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Height          =   735
      Left            =   -120
      TabIndex        =   36
      Top             =   1920
      Width           =   20775
   End
   Begin VB.Line Line2 
      X1              =   6120
      X2              =   5640
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exam"
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
      Left            =   6720
      TabIndex        =   29
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "REMARK"
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
      Index           =   12
      Left            =   11640
      TabIndex        =   14
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PERCENTAGE"
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
      Index           =   11
      Left            =   6840
      TabIndex        =   13
      Top             =   9960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
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
      Index           =   10
      Left            =   3120
      TabIndex        =   12
      Top             =   9960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "MARATHI"
      Height          =   495
      Index           =   9
      Left            =   11760
      TabIndex        =   11
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "HINDI"
      Height          =   495
      Index           =   8
      Left            =   7440
      TabIndex        =   10
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ENGLISH"
      Height          =   495
      Index           =   7
      Left            =   3000
      TabIndex        =   9
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "S.S"
      Height          =   495
      Index           =   6
      Left            =   11640
      TabIndex        =   8
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SCIENCE"
      Height          =   495
      Index           =   5
      Left            =   7440
      TabIndex        =   7
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "MATHS"
      Height          =   495
      Index           =   4
      Left            =   3120
      TabIndex        =   6
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "DIVISION"
      Height          =   495
      Index           =   3
      Left            =   15840
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "CLASS"
      Height          =   495
      Index           =   2
      Left            =   11760
      TabIndex        =   4
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      Height          =   495
      Index           =   1
      Left            =   7800
      TabIndex        =   3
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRN"
      Height          =   495
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   6360
      Width           =   975
   End
End
Attribute VB_Name = "RESULT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z As Integer
Dim sroll, sql, clas, divi, prnlogin, n1 As String
Dim outoftwo As Integer
Dim connect As New ADODB.Connection
Dim recset As New ADODB.Recordset




Private Sub back_Click()
For i = 0 To 12
Label1(i).Visible = False
overall.Enabled = True
sresult.Enabled = True
Next
For i = 0 To 12
Text1(i).Text = ""
Text1(i).Visible = False
Next
Combo1(0).Visible = False
Combo1(1).Visible = False
getresult.Visible = False
DataGrid1.Visible = False
Label2.Visible = False
examtype.Visible = False
End Sub





Private Sub Form_Load()
Set connect = New ADODB.Connection
connect.Open "newmethod", "system", "pict1228"
For i = 0 To 12
Label1(i).Visible = False
Label1(i).BackColor = &HE0E0E0
Next
For i = 0 To 12
'Text1(i).BackColor = &HFFFFFF
Text1(i).Text = ""
Text1(i).Visible = False
Next
Combo1(0).Visible = False
Combo1(1).Visible = False
getresult.Visible = False
DataGrid1.Visible = False
Label2.Visible = False
examtype.Visible = False
End Sub


Private Sub getresult_Click()
If examtype.Text = "exam type" Then
okay = MsgBox("Please select exam type", vbOKOnly)
Exit Sub
End If

If outoftwo = 1 Then    'indicates click on individual result
For i = 0 To 12
Label1(i).Visible = True
Next

For i = 0 To 12
Text1(i).Visible = True
Next
prnlogin = resultlogin.login

Set recset = New ADODB.Recordset
sql = "select r.prn,s.class,s.div,r.mat,r.sci,r.sst,r.eng,r.hin,r.mar from result r,sacademic s where r.testname='" + examtype.Text + "' and r.prn='" + prnlogin + "'"
Set recset = connect.Execute(sql)
connect.Execute "commit"
           If recset.EOF Then
           For i = 0 To 12
           Label1(i).Visible = False
           Text1(i).Visible = False
            Next
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If

Text1(0).Text = recset(0)
Text1(2).Text = recset(1)
Text1(3).Text = recset(2)
Text1(4).Text = recset(3)
Text1(5).Text = recset(4)
Text1(6).Text = recset(5)
Text1(7).Text = recset(6)
Text1(8).Text = recset(7)
Text1(9).Text = recset(8)
Text1(10).Text = Val(Text1(4).Text) + Val(Text1(5).Text) + Val(Text1(6).Text) + Val(Text1(7).Text) + Val(Text1(8).Text) + Val(Text1(9).Text)
If examtype.Text = "final" Then
Text1(11).Text = Val(Text1(10).Text) * 100 / 600
Else
Text1(11).Text = Val(Text1(10).Text) * 100 / 180
End If
Text1(12).Text = "PASS"
sql = "select name from spersonal where prn='" + prnlogin + "'"
Set recset = connect.Execute(sql)
           If recset.EOF Then
            MsgBox "Record not found!!!", vbOKOnly
            Exit Sub
            End If
Text1(1).Text = recset(0)

recset.close


ElseIf outoftwo = 0 Then    'indicates click on get overall result
If Combo1(0).Text = "" Or Combo1(1).Text = "" Then
okay = MsgBox("Please select class and division!", WARNING)
End If


If Combo1(0).Text <> "" And Combo1(1).Text <> "" Then

Set recset = New ADODB.Recordset

recset.CursorType = adOpenStatic
recset.CursorLocation = adUseClient
recset.LockType = adLockOptimistic

sql = "select a.roll,s.mat,s.sci,s.sst,s.eng,s.hin,s.mar from sacademic a inner join result s ON a.prn=s.prn where a.class='" & Combo1(0).Text & "' and a.div='" & Combo1(1).Text & "' and testname='" + examtype.Text + "'"
recset.Open sql, connect
If recset.EOF Then
 MsgBox "Record not found!", vbOKOnly
 Exit Sub
End If
DataGrid1.Visible = True
Set DataGrid1.DataSource = recset
DataGrid1.Refresh
DataGrid1.Font.Bold = True
DataGrid1.BorderStyle = dbgNoBorder
Label1(1).Caption = "Name"
Label1(0).Caption = "PRN"
Label1(0).Visible = False
Label1(1).Visible = False

End If

End If


End Sub


Private Sub Home1_Click()
menu.Show
Unload Me
End Sub

Private Sub overall_Click()
DataGrid1.Visible = False
sresult.Enabled = False
Combo1(0).Visible = True
Combo1(1).Visible = True
Combo1(0).List(0) = "5th"
Combo1(0).List(1) = "6th"
Combo1(0).List(2) = "7th"
Combo1(0).List(3) = "8th"
Combo1(0).List(4) = "9th"
Combo1(0).List(5) = "10th"
Combo1(0).List(6) = "11th"
Combo1(0).List(7) = "12th"

Combo1(1).List(0) = "A"
Combo1(1).List(1) = "B"
For z = 0 To 12
Text1(z).Visible = False
Label1(z).Visible = False
Next z

examtype.Visible = True
examtype.List(0) = "unittest1"
examtype.List(1) = "unittest2"
examtype.List(2) = "final"
For i = 0 To 2
examtype.Font.Bold = True
Next
Label2.Visible = True
For i = 0 To 12
Text1(i).Text = ""
Next
outoftwo = 0
Label1(0).Visible = True
Label1(1).Visible = True
Label1(0).Caption = "Class"
Label1(1).Caption = "Division"

getresult.Visible = True
End Sub

Private Sub sresult_Click()
overall.Enabled = False
DataGrid1.Visible = False
outoftwo = 1
resultlogin.Show
RESULT.Hide
examtype.Visible = True
examtype.List(0) = "unittest1"
examtype.List(1) = "unittest2"
examtype.List(2) = "final"
examtype.Text = "exam type"
For i = 0 To 2
examtype.Font.Bold = True
Next
getresult.Visible = True

End Sub


