VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Left            =   7680
      TabIndex        =   5
      Top             =   7680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   " OK"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   7680
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   8160
      TabIndex        =   3
      Text            =   " "
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   8160
      TabIndex        =   2
      Text            =   " "
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   " Password"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1920
      TabIndex        =   1
      Top             =   5160
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   " Admin"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1920
      TabIndex        =   0
      Top             =   2880
      Width           =   4815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If admin = Val(Text1.Text) And k = Val(Text2.Text) Then
MsgBox "login sucessfully"
Form4.Hide
Form5.Show
End If
End Sub

Private Sub Command2_Click()
End
End Sub
