VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   Caption         =   "LOGIN FORM"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20430
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   20430
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11280
      TabIndex        =   7
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9840
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NOT NOW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   1
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FORGET PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000018&
      BorderWidth     =   3
      Height          =   3615
      Left            =   6600
      Top             =   4440
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   8400
      Picture         =   "Form1.frx":0000
      Top             =   1200
      Width           =   2805
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text2.PasswordChar = ""
Else
Text2.PasswordChar = "*"
End If
End Sub

Private Sub Command2_Click()
Dim COUNT As Integer
While rs.EOF <> True
If rs(0) = Text1.Text And rs(1) = Text2.Text Then
COUNT = 1
End If
rs.MoveNext
Wend
If (COUNT = 1) Then
Unload Me
MDIForm1.Show
Else
If Text1.Text = "" Then
MsgBox "Username is required"

Else
If Text2.Text = "" Then
MsgBox "Password is required"

Else
MsgBox " WRONG USER OR PASSWORD"
End If
End If
End If
End Sub

Private Sub Command1_Click()
Unload Me
FORGETPASSWORDFORM.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "select * from Login", con, adOpenDynamic
rs.MoveFirst
End Sub

