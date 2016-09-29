VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   7080
      Top             =   9960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=itb24;Data Source=orcl;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=itb24;Data Source=orcl;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "itb24"
      Password        =   "itb24"
      RecordSource    =   "LIB"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1800
      Top             =   9960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=itb24;Data Source=orcl;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=itb24;Data Source=orcl;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "itb24"
      Password        =   "itb24"
      RecordSource    =   "PUPIL"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   10
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   9
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   8
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "ID_NO"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Text            =   " "
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      DataField       =   "ISSUE_DATE"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Text            =   " "
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "RENEWAL_DATE"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Text            =   " "
      Top             =   6720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Text            =   " "
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Renewal Date"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   3
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   2
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "ID No"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   1
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   2640
      Width           =   3135
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim con1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Private Sub Command1_Click()
If (Text3.Text = "") Then
MsgBox "enter issue date"
ElseIf (Text4.Text = "") Then
MsgBox "enter renewal date"
Else
MsgBox "updated"
con1.Execute "insert into lib values('" & Text3.Text & "','" & Text4.Text & "')"
con1.Execute "commit"
MsgBox "registration complete"
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End If
End Sub

Private Sub Command2_Click()
con.Execute "select * from pupil where name='" & Text1.Text & "'"
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
con.Open "Provider=MSDAORA.1;Password=itb24;User ID=itb24;Data Source=orcl;Persist Security Info=True"
rs.Open "select * from pupil", con, adOpenDynamic
con1.Open "Provider=MSDAORA.1;Password=itb24;User ID=itb24;Data Source=orcl;Persist Security Info=True"
rs1.Open "select * from lib", con, adOpenDynamic
End Sub

