VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16365
   LinkTopic       =   "Form10"
   ScaleHeight     =   10215
   ScaleWidth      =   16365
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "CLOSE"
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
      Left            =   7680
      TabIndex        =   18
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "ADD SUPPLIER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      Begin VB.CommandButton Command3 
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
         Height          =   375
         Left            =   4680
         TabIndex        =   17
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "NEW"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
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
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   6360
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   4320
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Text            =   "*****"
         Top             =   960
         Width           =   1815
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   855
         Left            =   600
         Top             =   6120
         Width           =   5535
      End
      Begin VB.Shape Shape7 
         Height          =   615
         Left            =   4560
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Shape Shape6 
         Height          =   615
         Left            =   2640
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         Height          =   615
         Left            =   720
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "DUES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "PINCODE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "ADDRESS"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "MOBILE_NO"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "COMPANY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER_ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   5055
         Left            =   480
         Top             =   720
         Width           =   5655
      End
   End
   Begin VB.Shape Shape4 
      Height          =   735
      Left            =   7560
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   9135
      Left            =   3960
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Text1.Text = "*****"
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Private Sub Command2_Click()
If rs.State = 1 Then rs.Close
rs.Open "SELECT  COUNT(S_ID) FROM SUPPLIER", con, adOpenDynamic
A = rs.Fields(0)
Text1.Text = "S" & "0" & (A + 1)
End Sub
Private Sub Command3_Click()
Dim STR As String
con.Execute "INSERT INTO SUPPLIER VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "')"
con.Execute "COMMIT"
MsgBox "ADDED SUCCESSFULLY"
Text1.Text = "*****"
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub



