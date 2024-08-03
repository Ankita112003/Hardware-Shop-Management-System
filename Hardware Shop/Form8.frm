VERSION 5.00
Begin VB.Form Form8 
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
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
      Height          =   615
      Left            =   9120
      TabIndex        =   19
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   18
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "FILTER OFF"
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
      Left            =   4800
      TabIndex        =   16
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FILTER ON"
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
      Left            =   7200
      TabIndex        =   15
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "UPDATE PRODUCTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   4560
      TabIndex        =   0
      Top             =   1800
      Width           =   9255
      Begin VB.CommandButton Command7 
         Caption         =   "SHOW"
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
         Left            =   4080
         TabIndex        =   17
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "NEXT"
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
         Left            =   5880
         TabIndex        =   14
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "FIRST"
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
         Left            =   480
         TabIndex        =   13
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "LAST"
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
         Left            =   7440
         TabIndex        =   12
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PREVIOUS"
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
         Left            =   2280
         TabIndex        =   11
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   5160
         TabIndex        =   10
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   5160
         TabIndex        =   9
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   5160
         TabIndex        =   8
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   5160
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   5160
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   3
         Height          =   975
         Left            =   240
         Top             =   5160
         Width           =   8775
      End
      Begin VB.Shape Shape6 
         Height          =   735
         Left            =   7320
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Shape Shape5 
         Height          =   735
         Left            =   5760
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Shape Shape4 
         Height          =   735
         Left            =   3960
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   2160
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   360
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "GST"
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
         Left            =   2760
         TabIndex        =   5
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "PRICE"
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
         Left            =   2880
         TabIndex        =   4
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "UNIT"
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
         Left            =   2880
         TabIndex        =   3
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "VARITIES"
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
         Left            =   2880
         TabIndex        =   2
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   3855
         Left            =   2280
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PRODUCT_NO"
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
         Left            =   3000
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Shape Shape14 
      BorderWidth     =   3
      Height          =   9975
      Left            =   3840
      Top             =   240
      Width           =   10695
   End
   Begin VB.Shape Shape13 
      BorderWidth     =   3
      Height          =   1095
      Left            =   6600
      Top             =   8880
      Width           =   4695
   End
   Begin VB.Shape Shape12 
      Height          =   855
      Left            =   9000
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Shape Shape11 
      Height          =   855
      Left            =   6720
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Shape Shape10 
      BorderWidth     =   3
      Height          =   975
      Left            =   4560
      Top             =   600
      Width           =   4695
   End
   Begin VB.Shape Shape9 
      Height          =   735
      Left            =   7080
      Top             =   720
      Width           =   1935
   End
   Begin VB.Shape Shape8 
      Height          =   735
      Left            =   4680
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim GST As Double
Private Sub Command1_Click()
rs.MovePrevious
If rs.BOF = True Then
rs.MoveFirst
MsgBox "FIRST USER"
Else
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End If
End Sub
Private Sub Command2_Click()
rs.MoveLast
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End Sub
Private Sub Command3_Click()
rs.MoveFirst
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End Sub

Private Sub Command4_Click()
rs.MoveNext
If rs.EOF = True Then
rs.MoveLast
MsgBox "LAST USER"
Else
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End If
End Sub

Private Sub Command5_Click()
Dim SEARCH1 As String
Dim SEARCH2 As String
SEARCH1 = InputBox("ENTER VARIETY NAME")
SEARCH2 = InputBox("ENTER UNIT")
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRODUCT where P_NAME LIKE '%" & SEARCH1 & "%' AND P_UNIT LIKE '%" & SEARCH2 & "%' ", con, adOpenDynamic
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End Sub

Private Sub Command6_Click()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRODUCT", con, adOpenDynamic
End Sub

Private Sub Command7_Click()
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
End Sub

Private Sub Command8_Click()
Dim STR As String
STR = "UPDATE PRODUCT SET P_RATE = '" & Text4.Text & "' , P_GST = '" & Text5.Text & "'  WHERE P_NO = '" & Text1.Text & "' "
con.Execute STR
con.Execute "COMMIT"
MsgBox "UPDATED SUCCESSFULLY"
End Sub

Private Sub Command9_Click()
Unload Me
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRODUCT", con, adOpenDynamic
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
GST = Val(Text4.Text) * 0.12
Text5.Text = GST
End Sub

