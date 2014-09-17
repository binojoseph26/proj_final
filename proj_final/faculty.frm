VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form faculty 
   Caption         =   "FACULTY"
   ClientHeight    =   7470
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11115
   LinkTopic       =   "Form2"
   ScaleHeight     =   7470
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   29
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
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
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
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Menu home 
      Caption         =   "Home"
   End
   Begin VB.Menu fac_info 
      Caption         =   "Faculty info"
      Begin VB.Menu secondary_fac 
         Caption         =   "secondary"
      End
      Begin VB.Menu high_fac 
         Caption         =   "Higher secondary"
      End
   End
   Begin VB.Menu class_fac 
      Caption         =   "Class faculty"
   End
End
Attribute VB_Name = "faculty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rss As New ADODB.Recordset
Dim cnn As New ADODB.Connection
Dim sql As String

Private Sub Form_Load()
Set cnn = New ADODB.Connection
cnn.Open "newmethod", "system", "pict1228"
DataGrid1.Visible = False
End Sub

Private Sub high_fac_Click()
DataGrid1.Visible = True
Label2.Caption = "Secondary department"
Label2.Visible = True
rss.CursorType = adOpenStatic
rss.CursorLocation = adUseClient
rss.LockType = adLockOptimistic
sql = "select t.id,t.name,t.email,p.qual,p.leaves from tpersonal t,toffice p where p.school='hsecondary' And t.id = p.id"
rss.Open sql, cnn
Set DataGrid1.DataSource = rss
End Sub

Private Sub home_Click()
menu.Show
faculty.Hide
End Sub

Private Sub secondary_fac_Click()
DataGrid1.Visible = True
Label2.Caption = "Secondary department"
Label2.Visible = True
rss.CursorType = adOpenStatic
rss.CursorLocation = adUseClient
rss.LockType = adLockOptimistic
sql = "select t.id,t.name,t.email,p.qual,p.leaves from tpersonal t,toffice p where p.school='secondary' And t.id = p.id"
rss.Open sql, cnn
Set DataGrid1.DataSource = rss
End Sub
