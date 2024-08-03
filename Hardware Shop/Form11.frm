VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form11 
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16395
   LinkTopic       =   "Form11"
   ScaleHeight     =   10515
   ScaleWidth      =   16395
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
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
      Height          =   495
      Left            =   9240
      TabIndex        =   21
      Top             =   10920
      Width           =   1455
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
      Height          =   495
      Left            =   7440
      TabIndex        =   20
      Top             =   10920
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form11.frx":0000
      Height          =   2415
      Left            =   3000
      TabIndex        =   19
      Top             =   7800
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   4260
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
      ColumnCount     =   7
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
         DataField       =   "S_NAME"
         Caption         =   "NAME"
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
         DataField       =   "S_COMP"
         Caption         =   "COMPANY"
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
         DataField       =   "S_MOBILE"
         Caption         =   "MOBILE_NO"
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
         DataField       =   "S_ADD"
         Caption         =   "ADDRESS"
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
         DataField       =   "S_ADD"
         Caption         =   "PINCODE"
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
         DataField       =   "S_DUES"
         Caption         =   "DUES"
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
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2055.118
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2039.811
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1934.929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5280
      Top             =   12360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM SUPPLIER"
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
      Caption         =   "SUPPLIER LIST"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   9855
      Begin VB.CommandButton Command5 
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
         Height          =   375
         Left            =   6960
         TabIndex        =   18
         Top             =   6240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
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
         Height          =   375
         Left            =   5400
         TabIndex        =   17
         Top             =   6240
         Width           =   1215
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
         Height          =   375
         Left            =   3600
         TabIndex        =   16
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "FIRST"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   6240
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         DataField       =   "S_DUES"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5520
         TabIndex        =   14
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         DataField       =   "S_PINCODE"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5520
         TabIndex        =   13
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         DataField       =   "S_ADD"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5520
         TabIndex        =   12
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         DataField       =   "S_MOBILE"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   525
         Left            =   5520
         TabIndex        =   11
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         DataField       =   "S_COMP"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   525
         Left            =   5520
         TabIndex        =   10
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         DataField       =   "S_NAME"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5520
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         DataField       =   "S_ID"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5520
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Shape Shape6 
         Height          =   615
         Left            =   6840
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Shape Shape5 
         Height          =   615
         Left            =   5280
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Shape Shape4 
         Height          =   615
         Left            =   3480
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         Height          =   615
         Left            =   2040
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   3
         Height          =   1095
         Left            =   1800
         Top             =   5880
         Width           =   6615
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   5175
         Left            =   2040
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
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
         Left            =   2400
         TabIndex        =   7
         Top             =   5040
         Width           =   855
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
         Left            =   2400
         TabIndex        =   6
         Top             =   4320
         Width           =   855
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
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   3600
         Width           =   975
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
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   2880
         Width           =   1215
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
         Left            =   2400
         TabIndex        =   3
         Top             =   2280
         Width           =   975
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
         Left            =   2400
         TabIndex        =   2
         Top             =   1680
         Width           =   855
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
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Shape Shape10 
      BorderWidth     =   3
      Height          =   11775
      Left            =   2160
      Top             =   120
      Width           =   13695
   End
   Begin VB.Shape Shape9 
      Height          =   735
      Left            =   9120
      Top             =   10800
      Width           =   1695
   End
   Begin VB.Shape Shape8 
      Height          =   735
      Left            =   7320
      Top             =   10800
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      BorderWidth     =   3
      Height          =   975
      Left            =   7200
      Top             =   10680
      Width           =   3735
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveFirst
End If
End Sub



Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command6_Click()
Unload Me
End Sub
