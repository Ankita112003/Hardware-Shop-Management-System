VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0E0FF&
   Caption         =   "SPLASH SCREEN"
   ClientHeight    =   9510
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   16380
   LinkTopic       =   "Form5"
   ScaleHeight     =   9510
   ScaleWidth      =   16380
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   35
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   4
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Height          =   255
      Left            =   8040
      TabIndex        =   3
      Top             =   6960
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   4455
      Left            =   7320
      Top             =   3480
      Width           =   7095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   10200
      TabIndex        =   2
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HARDWARE  SHOP  MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      TabIndex        =   1
      Top             =   4080
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YOUR APPLICATION IS LOADING ----"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   6000
      Width           =   4815
   End
   Begin VB.Image Image3 
      Height          =   4095
      Left            =   14760
      Picture         =   "Form5.frx":0000
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Image Image2 
      Height          =   2700
      Left            =   8640
      Picture         =   "Form5.frx":15778
      Top             =   480
      Width           =   4800
   End
   Begin VB.Image Image1 
      Height          =   3915
      Left            =   600
      Picture         =   "Form5.frx":174F0
      Top             =   3720
      Width           =   6255
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
con.ConnectionString = "Provider=MSDAORA.1;Password=AGARWAL;User ID=ANKITA;Persist Security Info=FALSE"
con.Open
Label1.Width = 0
Label2.Caption = 0 & " % "
Label3.Visible = False
Label2.Visible = False
End Sub



Private Sub Label5_Click()
Label5.Visible = False
Label3.Visible = True
Label2.Visible = True
End Sub
Private Sub Timer1_Timer()
If Label5.Visible = False Then
Label1.Width = Label1.Width + 56
Label2.Caption = Val(Label2.Caption) + 1 & " % "
If Val(Label2.Caption) > 100 Then
Timer1.Enabled = False
Form1.Show
Unload Me
Else
End If
End If
End Sub

