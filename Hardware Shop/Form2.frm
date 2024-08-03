VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000018&
   Caption         =   "CREATE USER"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   7
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9480
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9480
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      Height          =   735
      Left            =   9120
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   6840
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   975
      Left            =   6720
      Top             =   5640
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000018&
      BorderColor     =   &H80000012&
      BorderWidth     =   3
      Height          =   4935
      Left            =   5040
      Top             =   2040
      Width           =   8055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER USER_ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER MOBILE_NO."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim STR As String
STR = "insert into login values ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "')"
con.Execute STR
MsgBox "NEW USER ADDED SUCESSFULLY"
con.Execute "commit"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

