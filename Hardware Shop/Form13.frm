VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form13 
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form13"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   240
      TabIndex        =   63
      Top             =   10080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   201588737
      CurrentDate     =   45339
   End
   Begin VB.TextBox Text18 
      Height          =   615
      Left            =   360
      TabIndex        =   62
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   19800
      TabIndex        =   61
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   19800
      TabIndex        =   60
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18360
      TabIndex        =   59
      Top             =   11160
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SELL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18360
      TabIndex        =   58
      Top             =   10080
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      Height          =   495
      Left            =   13560
      TabIndex        =   56
      Top             =   10800
      Width           =   2055
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   495
      Left            =   13560
      TabIndex        =   55
      Top             =   9840
      Width           =   2055
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   6360
      TabIndex        =   54
      Top             =   10920
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      TabIndex        =   53
      Top             =   9840
      Width           =   2055
   End
   Begin VB.ListBox List12 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":0000
      Left            =   16200
      List            =   "Form13.frx":0002
      TabIndex        =   48
      Top             =   6120
      Width           =   2055
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":0004
      Left            =   14880
      List            =   "Form13.frx":0006
      TabIndex        =   46
      Top             =   6120
      Width           =   1335
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":0008
      Left            =   13680
      List            =   "Form13.frx":000A
      TabIndex        =   45
      Top             =   6120
      Width           =   1215
   End
   Begin VB.ListBox List11 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":000C
      Left            =   11760
      List            =   "Form13.frx":000E
      TabIndex        =   41
      Top             =   6120
      Width           =   1935
   End
   Begin VB.ListBox List10 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":0010
      Left            =   9720
      List            =   "Form13.frx":0012
      TabIndex        =   39
      Top             =   6120
      Width           =   2055
   End
   Begin VB.ListBox List9 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":0014
      Left            =   7920
      List            =   "Form13.frx":0016
      TabIndex        =   37
      Top             =   6120
      Width           =   1815
   End
   Begin VB.ListBox List8 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":0018
      Left            =   6120
      List            =   "Form13.frx":001A
      TabIndex        =   35
      Top             =   6120
      Width           =   1815
   End
   Begin VB.ListBox List7 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":001C
      Left            =   4440
      List            =   "Form13.frx":001E
      TabIndex        =   33
      Top             =   6120
      Width           =   1695
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":0020
      Left            =   2760
      List            =   "Form13.frx":0022
      TabIndex        =   31
      Top             =   6120
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "Form13.frx":0024
      Left            =   1800
      List            =   "Form13.frx":0026
      TabIndex        =   30
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   375
      Left            =   17280
      TabIndex        =   29
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   375
      Left            =   15360
      TabIndex        =   28
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13680
      TabIndex        =   27
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Left            =   11880
      TabIndex        =   26
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   9720
      TabIndex        =   25
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   23
      Top             =   4680
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form13.frx":0028
      Left            =   3480
      List            =   "Form13.frx":004D
      TabIndex        =   22
      Text            =   "CHOOSE UNIT"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form13.frx":0098
      Left            =   1320
      List            =   "Form13.frx":00B1
      TabIndex        =   21
      Text            =   "SELECT VARIETIES"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   12120
      TabIndex        =   20
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   12120
      TabIndex        =   19
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7680
      TabIndex        =   18
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   17
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape9 
      BorderWidth     =   3
      Height          =   2055
      Left            =   19560
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Shape Shape8 
      Height          =   855
      Left            =   19680
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Shape Shape7 
      Height          =   855
      Left            =   19680
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Shape Shape6 
      BorderWidth     =   3
      Height          =   2175
      Left            =   18120
      Top             =   9840
      Width           =   2175
   End
   Begin VB.Shape Shape5 
      Height          =   855
      Left            =   18240
      Top             =   11040
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      Height          =   855
      Left            =   18240
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   "CALCULATIONS "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   57
      Top             =   9240
      Width           =   2775
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   2295
      Left            =   2880
      Top             =   9360
      Width           =   13695
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL WITH TAX"
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
      Left            =   10680
      TabIndex        =   52
      Top             =   10920
      Width           =   2295
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL AMOUNT"
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
      Left            =   10680
      TabIndex        =   51
      Top             =   9960
      Width           =   2175
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT MODE"
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
      Left            =   3240
      TabIndex        =   50
      Top             =   10920
      Width           =   2535
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO OF PRODUCTS"
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
      Left            =   3240
      TabIndex        =   49
      Top             =   9960
      Width           =   2535
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL PRICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16200
      TabIndex        =   47
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AMOUNT"
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
      Left            =   14880
      TabIndex        =   44
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
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
      Left            =   13680
      TabIndex        =   43
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   13680
      TabIndex        =   42
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   735
      Left            =   11760
      TabIndex        =   40
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      TabIndex        =   38
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   36
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   735
      Left            =   6120
      TabIndex        =   34
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   735
      Left            =   4440
      TabIndex        =   32
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRODUCT NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   16
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SERIAL NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   15
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL PRICE"
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
      Left            =   17280
      TabIndex        =   14
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GST PRICE"
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
      Left            =   15240
      TabIndex        =   13
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   13680
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   11880
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AVL_QUANTITY"
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
      Left            =   9600
      TabIndex        =   10
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
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
      Left            =   7560
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   1440
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "ADD PRODUCTS"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SELL _NO"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SELL DATE"
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
      Left            =   10320
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NO"
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
      Left            =   10440
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER NAME"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label24 
      Caption         =   "   MAIN INFORMATION"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   5295
      Left            =   1080
      Top             =   3600
      Width           =   17775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   2415
      Left            =   4440
      Top             =   480
      Width           =   10575
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W As Integer
Dim ZZ As String
Dim KK As Integer
Dim S As Double
Dim P As String
Dim C As Integer




Private Sub Combo1_Click()
If Combo2.Text = "Half Pound(226 gm)" Or Combo2.Text = "Full Pound(453 gm)" Or Combo2.Text = "Two Pound(907 gm)" Then
If rs.State = 1 Then rs.Close
rs.Open "select * from PRODUCT WHERE P_NAME= '" & Combo1.Text & "' AND P_UNIT ='" & Combo2.Text & "' ", con, adOpenDynamic
Text5.Text = rs.Fields(0)

End If

End Sub

Private Sub Combo2_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from PRODUCT WHERE P_NAME= '" & Combo1.Text & "' AND P_UNIT ='" & Combo2.Text & "' ", con, adOpenDynamic
Text5.Text = rs.Fields(0)
If rs.State = 1 Then rs.Close
rs.Open "select * from STOCK WHERE P_NAME= '" & Combo1.Text & "' AND P_UNIT ='" & Combo2.Text & "' ", con, adOpenDynamic
Text17.Text = rs.Fields(3)

End Sub
Private Sub Command1_Click()
Text14.Text = (Val(Text14.Text) - Val(List11.List(List11.ListCount - 1)))
Text13.Text = (Val(Text13.Text) - Val(List12.List(List12.ListCount - 1)))
List1.RemoveItem (List1.ListCount - 1)
List4.RemoveItem (List4.ListCount - 1)
List5.RemoveItem (List5.ListCount - 1)
List6.RemoveItem (List6.ListCount - 1)
List7.RemoveItem (List7.ListCount - 1)
List8.RemoveItem (List8.ListCount - 1)
List9.RemoveItem (List9.ListCount - 1)
List10.RemoveItem (List10.ListCount - 1)
List11.RemoveItem (List11.ListCount - 1)
List12.RemoveItem (List12.ListCount - 1)
Text10.Text = List1.ListCount
C = C - 1
End Sub

Private Sub Command2_Click()

C = C + 1
List1.AddItem (C)
List6.AddItem Text5.Text
List7.AddItem Combo1.Text
List8.AddItem Combo2.Text
List4.AddItem Text7.Text
List5.AddItem Text9.Text
List10.AddItem Text11.Text
List11.AddItem Text6.Text
List12.AddItem Text8.Text
List9.AddItem (Val(Text6.Text) / Val(Text11.Text))
Text10.Text = List1.ListCount
Combo1.Text = "SELECT VARIETIES"
Combo2.Text = "CHOOSE UNIT"
Text11.Text = ""
Text5.Text = ""
A = Val(Text6.Text)
b = Val(Text8.Text)
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text17.Text = ""
Text14.Text = A + Val(Text14.Text)
Text13.Text = b + Val(Text13.Text)
End Sub

Private Sub Command3_Click()
Dim STR As String
STR = "INSERT INTO Sell_MASTER VALUES('" & Text1.Text & "','" & Text4.Text & "','" & Text2.Text & "','" & Text3.Text & "'," & Val(Text10.Text) & ",'" & Text12.Text & "'," & Val(Text14.Text) & "," & Val(Text13.Text) & ",'" & Text18.Text & "')"
con.Execute STR
con.Execute "COMMIT"

For I = 0 To List6.ListCount - 1
con.Execute "INSERT INTO Sell_DETAILS VALUES('" & Text1.Text & "','" & List6.List(I) & "'," & Val(List10.List(I)) & "," & Val(List12.List(I)) & ",'" & List7.List(I) & "'," & Val(List9.List(I)) & ",'" & Val(List4.List(I)) & "','" & List8.List(I) & "')"
con.Execute "COMMIT"
Next

Dim STR2 As String
For I = 0 To List6.ListCount - 1
STR2 = "UPDATE STOCK SET AVL_QUANTITY =  " & Val(List10.List(I)) & " WHERE P_NO = '" & List6.List(I) & "' "
con.Execute STR2
con.Execute "COMMIT"
Next


MsgBox "SELL COMPLETED"
STR = Trim(Text1.Text)
If DataEnvironment1.rsCommand4.State = 1 Then DataEnvironment1.rsCommand4.Close
DataEnvironment1.Command4 STR
DataReport4.Show

Unload Me
Load Me
End Sub

Private Sub Command4_Click()
Unload Me
Load Me
End Sub

Private Sub Command5_Click()
Text14.Text = (Val(Text14.Text) - Val(List11.List(Val(Text15.Text) - 1)))
Text13.Text = (Val(Text13.Text) - Val(List12.List(Val(Text15.Text) - 1)))
List1.RemoveItem (Val(Text15.Text) - 1)
List4.RemoveItem (Val(Text15.Text) - 1)
List5.RemoveItem (Val(Text15.Text) - 1)
List6.RemoveItem (Val(Text15.Text) - 1)
List7.RemoveItem (Val(Text15.Text) - 1)
List8.RemoveItem (Val(Text15.Text) - 1)
List9.RemoveItem (Val(Text15.Text) - 1)
List10.RemoveItem (Val(Text15.Text) - 1)
List11.RemoveItem (Val(Text15.Text) - 1)
List12.RemoveItem (Val(Text15.Text) - 1)
Text10.Text = List1.ListCount
C = C - 1
For I = (Val(Text15.Text) - 1) To (List1.ListCount - 1)
List1.List(I) = (Val(List1.List(I)) - 1)
Next
End Sub

Private Sub Command6_Click()
K = Val(List11.List(Val(Text15.Text) - 1)) / Val(List10.List(Val(Text15.Text) - 1))
J = Val(List12.List(Val(Text15.Text) - 1)) / Val(List10.List(Val(Text15.Text) - 1))
List10.RemoveItem (Val(Text15.Text) - 1)
List10.List(Val(Text15.Text) - 1) = Val(Text16.Text)
List11.List(Val(Text15.Text) - 1) = (K * (Val(List11.List(Val(Text15.Text) - 1))))
List12.List(Val(Text15.Text) - 1) = (J * (Val(List12.List(Val(Text15.Text) - 1))))

End Sub

Private Sub Form_Activate()
C = 0
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM Sell_MASTER", con, adOpenDynamic
If rs.BOF = True Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT count(S_NO) FROM Sell_MASTER", con, adOpenDynamic
A = rs.Fields(0)
Text1.Text = "S" & "O" & "0" & (A + 1)
Text18.Text = "I" & "N" & "V" & "0" & "0" & (A + 1)

P = Text1.Text
ElseIf P >= "SO09" Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(S_NO) FROM Sell_MASTER", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "S" & "O" & (A + 1)
Text18.Text = "I" & "N" & "V" & "0" & "0" & (A + 1)

P = Text1.Text
Else
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(S_NO) FROM Sell_MASTER", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "S" & "O" & "0" & (A + 1)
Text18.Text = "I" & "N" & "V" & "0" & "0" & (A + 1)

P = Text1.Text
End If
Text4.Text = Format(Date, "DD-MMM-YYYY")

Text18.Text = "I" & "N" & "V" & "0" & "0" & (A + 1)







End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If rs.State = 1 Then rs.Close
rs.Open "select * from PRODUCT WHERE P_NAME= '" & Combo1.Text & "' AND P_UNIT ='" & Combo2.Text & "' ", con, adOpenDynamic
Text6.Text = rs.Fields(3) * Val(Text11.Text)
Text7.Text = rs.Fields(4) & "%"
Text9.Text = rs.Fields(5) * Val(Text11.Text)
Text8.Text = rs.Fields(6) * Val(Text11.Text)
End Sub


