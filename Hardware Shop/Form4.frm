VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000018&
   Caption         =   "UPDATE USER"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   12495
   ScaleWidth      =   22920
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
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
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
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
      Left            =   9480
      TabIndex        =   3
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   3
      Height          =   1095
      Left            =   6960
      Top             =   6600
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      Height          =   855
      Left            =   9360
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      Height          =   855
      Left            =   7080
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3615
      Left            =   5400
      Top             =   2400
      Width           =   7215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER  REGISTERED USER_ID"
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
      Left            =   6360
      TabIndex        =   5
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER NEW MOBILE_NO."
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
      Left            =   6600
      TabIndex        =   2
      Top             =   4680
      Width           =   2415
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
rs.MovePrevious
If rs.BOF = True Then
rs.MoveFirst
MsgBox "FIRST USER"
Else
Text1.Text = rs(0)
Text3.Text = rs(2)
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = rs(0) Then
Dim STR As String
STR = "UPDATE LOGIN SET MOBILE = '" & Text3.Text & "' WHERE USERID = '" & Text1.Text & "' "
con.Execute STR
con.Execute "COMMIT"
MsgBox "UPDATED SUCCESSFULLY"
Else
MsgBox "USER_ID CAN NOT BE UPDATED"
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Text1.Text = rs(0)
Text3.Text = rs(2)
End Sub

Private Sub Command5_Click()
rs.MoveNext
If rs.EOF = True Then
rs.MoveLast
MsgBox "LAST USER"
Else
Text1.Text = rs(0)
Text3.Text = rs(2)
End If
End Sub

Private Sub Form_Load()
rs.MoveFirst
End Sub

