VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   4800
      Top             =   11040
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
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
      Caption         =   "Back"
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
      Left            =   12000
      TabIndex        =   12
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
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
      Left            =   12000
      TabIndex        =   11
      Top             =   5160
      Width           =   2415
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
      Height          =   735
      Left            =   12000
      TabIndex        =   10
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      DataField       =   "MOBILE_NO"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   7560
      TabIndex        =   9
      Text            =   " "
      Top             =   7800
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      DataField       =   "BOOK_NAME"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   7560
      TabIndex        =   8
      Text            =   " "
      Top             =   6240
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "ID_NO"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   7560
      TabIndex        =   7
      Text            =   " "
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "AGE"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   7560
      TabIndex        =   6
      Text            =   " "
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   7560
      TabIndex        =   5
      Text            =   " "
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Mobile No"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   4
      Top             =   7680
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "Bookname"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   6360
      Width           =   4695
   End
   Begin VB.Label Label3 
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
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   4920
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Age"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   3240
      Width           =   4455
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
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   1800
      Width           =   4695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
If (Text1.Text = "") Then
MsgBox "enter the name"
ElseIf (Text2.Text = "") Then
MsgBox "enter the age"
ElseIf (Text3.Text = "") Then
MsgBox "enter idno"
ElseIf (Text4.Text = "") Then
MsgBox "enter bookname"
ElseIf (Text5.Text = "") Then
MsgBox "enter mobile no"
Else
MsgBox "Updated"
con.Execute "insert into pupil values('" & Text1.Text & "','" & Text2.Text & "', '" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "')"
con.Execute "commit"
MsgBox "registration complete"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command3_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Form_Load()
con.Open "Provider=MSDAORA.1;Password=itb24;User ID=itb24;Data Source=orcl;Persist Security Info=True"
rs.Open "select * from pupil", con, adOpenDynamic
End Sub

