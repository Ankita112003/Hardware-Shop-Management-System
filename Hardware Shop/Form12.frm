VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form12 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16965
   LinkTopic       =   "Form12"
   ScaleHeight     =   10815
   ScaleWidth      =   16965
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   19
      Top             =   10080
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   10080
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form12.frx":0000
      Height          =   2655
      Left            =   3840
      TabIndex        =   17
      Top             =   6960
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4683
      _Version        =   393216
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "S_ID"
         Caption         =   "SUPPLIER_ID"
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
         DataField       =   "P_NO"
         Caption         =   "PRODUCT_NO"
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
         DataField       =   "S_RATE"
         Caption         =   "RATE"
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
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3030.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3390.236
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   17880
      Top             =   12240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
      Enabled         =   0
      Connect         =   "Provider=MSDAORA.1;User ID=ANKITA/AGARWAL;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=ANKITA/AGARWAL;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM SUPPLIER_PRODUCT ORDER BY S_ID"
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
      BackColor       =   &H8000000B&
      Caption         =   "SUPPLIER PRODUCT PRICE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   5520
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.CommandButton Command4 
         Caption         =   "DELETE"
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
         Left            =   4560
         TabIndex        =   16
         Top             =   5640
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   5640
         Width           =   855
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
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   5640
         Width           =   855
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
         Left            =   600
         TabIndex        =   13
         Top             =   5640
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form12.frx":0015
         Left            =   3120
         List            =   "Form12.frx":0037
         TabIndex        =   12
         Text            =   "CHOOSE UNIT"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form12.frx":0078
         Left            =   3120
         List            =   "Form12.frx":0091
         TabIndex        =   11
         Text            =   "SELECT VARIETY"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Text            =   "SELECT SUPPLIER"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   4560
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Text            =   "*****"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Text            =   "*****"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Shape Shape7 
         Height          =   615
         Left            =   4440
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Shape Shape6 
         Height          =   615
         Left            =   3240
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         Height          =   615
         Left            =   1680
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         Height          =   615
         Left            =   480
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   3
         Height          =   855
         Left            =   3120
         Top             =   5400
         Width           =   2535
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   3
         Height          =   855
         Left            =   360
         Top             =   5400
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         Left            =   960
         TabIndex        =   5
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Left            =   960
         TabIndex        =   4
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "VARIETY"
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
         Left            =   960
         TabIndex        =   3
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Left            =   960
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER NAME"
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
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   4695
         Left            =   480
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Shape Shape11 
      BorderWidth     =   3
      Height          =   10455
      Left            =   1800
      Top             =   240
      Width           =   13815
   End
   Begin VB.Shape Shape10 
      Height          =   615
      Left            =   8880
      Top             =   9960
      Width           =   1215
   End
   Begin VB.Shape Shape9 
      Height          =   615
      Left            =   7440
      Top             =   9960
      Width           =   1215
   End
   Begin VB.Shape Shape8 
      BorderWidth     =   3
      Height          =   855
      Left            =   7320
      Top             =   9840
      Width           =   2895
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
If rs.State = 1 Then rs.Close
rs.Open "select S_ID from SUPPLIER WHERE S_NAME= '" & Combo1.Text & "' ", con, adOpenDynamic
Text1.Text = rs.Fields(0)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_Click()
If rs.State = 1 Then rs.Close
rs.Open "select P_NO from PRODUCT WHERE P_NAME= '" & Combo2.Text & "' AND P_UNIT ='" & Combo3.Text & "' ", con, adOpenDynamic
Text2.Text = rs.Fields(0)
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Combo1.Text = "SELECT SUPPLIER"
Text1.Text = "*****"
Combo2.Text = "SELECT VARIETY"
Combo3.Text = "CHOOSE UNIT"
Text2.Text = "*****"
Text3.Text = ""
End Sub

Private Sub Command2_Click()
Dim STR As String
STR = "INSERT INTO SUPPLIER_PRODUCT VALUES('" & Text1.Text & "','" & Text2.Text & "'," & Val(Text3.Text) & ")"
con.Execute STR
con.Execute "COMMIT"
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Dim STR As String
STR = "UPDATE SUPPLIER_PRODUCT SET RATE = " & Val(Text3.Text) & " WHERE S_ID ='" & Text1.Text & "' AND P_NO ='" & Text2.Text & "' "
con.Execute STR
con.Execute "COMMIT"
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
Dim STR As String
STR = "DELETE FROM SUPPLIER_PRODUCT WHERE S_ID ='" & Text1.Text & "' AND P_NO ='" & Text2.Text & "' "
con.Execute STR
con.Execute "COMMIT"
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
Unload Me
End Sub


Private Sub Form_Activate()
If rs.State = 1 Then rs.Close
rs.Open "select * from SUPPLIER ", con, adOpenDynamic
While rs.EOF = False
Combo1.AddItem rs.Fields(1)
rs.MoveNext
Wend
End Sub





