VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form6.frx":0000
      Height          =   3495
      Left            =   3960
      TabIndex        =   18
      Top             =   8760
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "P_NO"
         Caption         =   "ProductNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "P_NAME"
         Caption         =   "Varieties"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "P_UNIT"
         Caption         =   "Unit"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "P_RATE"
         Caption         =   "Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "P_GST"
         Caption         =   "GST(%)"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "GST_AMOUNT"
         Caption         =   "Gst Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "S_RATE"
         Caption         =   "Selling Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   17880
      Top             =   12360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=MSDAORA.1;User ID=ANKITA/AGARWAL;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=ANKITA/AGARWAL;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "ANKITA/AGARWAL"
      Password        =   ""
      RecordSource    =   "SELECT * FROM PRODUCT ORDER BY P_NO"
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
   Begin VB.Frame Frame1 
      Caption         =   "ADD PRODUCTS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   5520
      TabIndex        =   0
      Top             =   720
      Width           =   6975
      Begin VB.CommandButton Command4 
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
         Left            =   4560
         TabIndex        =   17
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Height          =   615
         Left            =   2880
         TabIndex        =   16
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
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
         Height          =   615
         Left            =   1080
         TabIndex        =   15
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3480
         TabIndex        =   14
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   3480
         TabIndex        =   13
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3480
         TabIndex        =   11
         Top             =   3120
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form6.frx":0015
         Left            =   3480
         List            =   "Form6.frx":0037
         TabIndex        =   10
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form6.frx":0079
         Left            =   3480
         List            =   "Form6.frx":0092
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3480
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   3
         Height          =   1095
         Left            =   720
         Top             =   6120
         Width           =   5415
      End
      Begin VB.Shape Shape4 
         Height          =   855
         Left            =   4440
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         Height          =   855
         Left            =   2760
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         Height          =   855
         Left            =   960
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   5055
         Left            =   840
         Top             =   720
         Width           =   5295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "SELLING PRICE"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "GST AMOUNT"
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
         Left            =   1440
         TabIndex        =   6
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "GST(%)"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   3840
         Width           =   1215
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
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   3240
         Width           =   975
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
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "VARIETIES"
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
         Left            =   1560
         TabIndex        =   2
         Top             =   1800
         Width           =   1095
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
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Shape Shape6 
      BorderWidth     =   3
      Height          =   7935
      Left            =   5160
      Top             =   480
      Width           =   7695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GST As Double
Dim P As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Command2_Click()
Dim STR As String
STR = "INSERT INTO PRODUCT VALUES('" & Text1.Text & "','" & Combo1.Text & "','" & Combo2.Text & "'," & Val(Text2.Text) & "," & Val(Text4.Text) & "," & Val(Text3.Text) & "," & Val(Text5.Text) & ")"
con.Execute STR
con.Execute "INSERT INTO STOCK VALUES('" & Text1.Text & "','" & Combo1.Text & "','" & Combo2.Text & "'," & 0 & ")"
con.Execute "COMMIT"
MsgBox "ADDED SUCCESSFULLY"
Adodc1.Refresh
Unload Me
Load Me
End Sub

Private Sub Command3_Click()
Text1.Text = "*****"
Combo1.Text = "SELECT"
Combo2.Text = "CHOOSE"
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRODUCT ", con, adOpenDynamic
If rs.BOF = True Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT count(p_no) FROM PRODUCT", con, adOpenDynamic
A = rs.Fields(0)
Text1.Text = "P" & "0" & (A + 1)
P = Text1.Text
ElseIf P >= "P09" Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(p_no) FROM PRODUCT", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "P" & (A + 1)
P = Text1.Text
Else
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(p_no) FROM PRODUCT", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "P" & "0" & (A + 1)
P = Text1.Text
End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
GST = (Val(Text4.Text) / 100) * Val(Text2.Text)
Text3.Text = GST
Text5.Text = Val(Text2.Text) + Val(Text3.Text)
End Sub


