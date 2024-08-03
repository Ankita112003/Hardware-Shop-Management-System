VERSION 5.00
Begin VB.Form RESETPASSWORDFORM 
   Caption         =   "RESETPASSWORDFORM"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11520
      TabIndex        =   5
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET PASSWORD"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   4
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9960
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3975
      Left            =   6600
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONFORM PASSWORD"
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
      Left            =   6840
      TabIndex        =   1
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NEW PASSWORD"
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
      Left            =   7200
      TabIndex        =   0
      Top             =   3840
      Width           =   1815
   End
End
Attribute VB_Name = "RESETPASSWORDFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.PasswordChar = ""
Else
Text1.PasswordChar = "*"
End If
End Sub

Private Sub Command2_Click()
If Len(Text1.Text) < 8 Then
MsgBox "MAKE ATLEAST 6 CHARACTERS PASSWORD"
ElseIf Text1.Text <> Text2.Text Then
MsgBox "PASSWORD NOT MATCHED ENTER CAREFULLY"
Else
Dim STR As String
STR = "UPDATE Login set password = '" & Text2.Text & "' WHERE USERID = '" & PASS & "' "
con.Execute STR
con.Execute "commit"
Unload Me
Form1.Show
End If
End Sub

