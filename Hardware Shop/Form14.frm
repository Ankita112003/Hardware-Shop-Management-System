VERSION 5.00
Begin VB.Form Form14 
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form14"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   19320
      TabIndex        =   60
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
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
      Left            =   19320
      TabIndex        =   59
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PURCHASE"
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
      Left            =   19320
      TabIndex        =   58
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
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
      Left            =   19320
      TabIndex        =   57
      Top             =   7560
      Width           =   1815
   End
   Begin VB.TextBox Text18 
      Height          =   495
      Left            =   15600
      TabIndex        =   56
      Top             =   10920
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8640
      TabIndex        =   55
      Top             =   11280
      Width           =   1935
   End
   Begin VB.TextBox Text16 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8640
      TabIndex        =   54
      Top             =   10440
      Width           =   1935
   End
   Begin VB.TextBox Text19 
      Height          =   495
      Left            =   12120
      TabIndex        =   53
      Top             =   11280
      Width           =   1815
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   52
      Top             =   11400
      Width           =   1815
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   12120
      TabIndex        =   51
      Top             =   10440
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   50
      Top             =   10440
      Width           =   1815
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
      Height          =   2310
      ItemData        =   "Form14.frx":0000
      Left            =   12600
      List            =   "Form14.frx":0002
      TabIndex        =   40
      Top             =   6960
      Width           =   1815
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
      Height          =   2310
      ItemData        =   "Form14.frx":0004
      Left            =   10800
      List            =   "Form14.frx":0006
      TabIndex        =   39
      Top             =   6960
      Width           =   1815
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
      Height          =   2310
      ItemData        =   "Form14.frx":0008
      Left            =   9120
      List            =   "Form14.frx":000A
      TabIndex        =   38
      Top             =   6960
      Width           =   1695
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      ItemData        =   "Form14.frx":000C
      Left            =   7080
      List            =   "Form14.frx":000E
      TabIndex        =   37
      Top             =   6960
      Width           =   2055
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      ItemData        =   "Form14.frx":0010
      Left            =   5040
      List            =   "Form14.frx":0012
      TabIndex        =   36
      Top             =   6960
      Width           =   2055
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
      Height          =   2310
      ItemData        =   "Form14.frx":0014
      Left            =   4320
      List            =   "Form14.frx":0016
      TabIndex        =   35
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   12480
      TabIndex        =   28
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   495
      Left            =   9840
      TabIndex        =   27
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   495
      Left            =   14880
      TabIndex        =   26
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   25
      Top             =   5160
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form14.frx":0018
      Left            =   4440
      List            =   "Form14.frx":003D
      TabIndex        =   24
      Text            =   "CHOOSE UNIT"
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form14.frx":0088
      Left            =   1680
      List            =   "Form14.frx":00A1
      TabIndex        =   23
      Text            =   "SELECT VARITIES"
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   495
      Left            =   15240
      TabIndex        =   14
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   495
      Left            =   15240
      TabIndex        =   13
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9240
      TabIndex        =   12
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   9480
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3960
      TabIndex        =   10
      Text            =   "SELECT SUPPLIER"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   15240
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.Shape Shape9 
      BorderWidth     =   3
      Height          =   2295
      Left            =   19080
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Shape Shape8 
      Height          =   855
      Left            =   19200
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Shape Shape7 
      Height          =   975
      Left            =   19200
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Shape Shape6 
      BorderWidth     =   3
      Height          =   2415
      Left            =   19080
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Shape Shape5 
      Height          =   975
      Left            =   19200
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      Height          =   855
      Left            =   19200
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "CALCULATIONS"
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
      Left            =   1920
      TabIndex        =   49
      Top             =   9720
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   2295
      Left            =   1680
      Top             =   9840
      Width           =   15855
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAID"
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
      Left            =   11160
      TabIndex        =   48
      Top             =   11400
      Width           =   615
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NET AMOUNT"
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
      Left            =   7080
      TabIndex        =   47
      Top             =   11400
      Width           =   1335
   End
   Begin VB.Label Label27 
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
      Left            =   2400
      TabIndex        =   46
      Top             =   11520
      Width           =   1695
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   14280
      TabIndex        =   45
      Top             =   11040
      Width           =   735
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAY"
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
      Left            =   11040
      TabIndex        =   44
      Top             =   10560
      Width           =   855
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PREV DUES"
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
      Left            =   7080
      TabIndex        =   43
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO. OF PRODUCTS"
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
      TabIndex        =   42
      Top             =   10560
      Width           =   1815
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "ADD PRODUCTS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   41
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   5415
      Left            =   1320
      Top             =   4080
      Width           =   16095
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   12600
      TabIndex        =   34
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   10800
      TabIndex        =   33
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   9120
      TabIndex        =   32
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   7080
      TabIndex        =   31
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRODUCT NO."
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
      Left            =   5040
      TabIndex        =   30
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SERIAL NO."
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
      Left            =   4320
      TabIndex        =   29
      Top             =   6360
      Width           =   735
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
      Left            =   15000
      TabIndex        =   22
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label14 
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
      Height          =   375
      Left            =   12240
      TabIndex        =   21
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label13 
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
      Left            =   9600
      TabIndex        =   20
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT NO."
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
      Left            =   6720
      TabIndex        =   19
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label11 
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
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label10 
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
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "MAIN INFORMATION"
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
      Left            =   1920
      TabIndex        =   16
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3375
      Left            =   1200
      Top             =   240
      Width           =   16335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   12480
      TabIndex        =   7
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY NAME"
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
      Left            =   12480
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER DATE"
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
      Left            =   12480
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   7080
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NO."
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE_NO."
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
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As Integer

Private Sub Combo2_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from PRODUCT WHERE P_NAME= '" & Combo1.Text & "' AND P_UNIT ='" & Combo2.Text & "' ", con, adOpenDynamic
Text8.Text = rs.Fields(0)
If rs.State = 1 Then rs.Close
rs.Open "select * from SUPPLIER_PRODUCT WHERE S_ID= '" & Text3.Text & "' AND P_NO ='" & Text8.Text & "' ", con, adOpenDynamic
Text12.Text = rs.Fields(2)


End Sub

Private Sub Combo3_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from SUPPLIER WHERE S_NAME='" & Combo3.Text & "' ", con, adOpenDynamic
Text3.Text = rs.Fields(0)
Text7.Text = rs.Fields(3)
Text5.Text = rs.Fields(2)
Text4.Text = rs.Fields(4)
Text2.Text = rs.Fields(5)
Text16.Text = rs.Fields(6)
End Sub

Private Sub Command1_Click()
Unload Me
Load Me

End Sub

Private Sub Command2_Click()
C = C + 1
List1.AddItem (C)
List2.AddItem Text8.Text
List3.AddItem Combo1.Text
List4.AddItem Combo2.Text
List5.AddItem Text11.Text
List6.AddItem Text9.Text
A = Val(Text9.Text)
Text8.Text = ""
Text9.Text = ""
Text11.Text = ""
Text12.Text = ""
Combo1.Text = "SELECT VARIETIES"
Combo2.Text = "CHOOSE UNIT"
Text10.Text = List1.ListCount
Text14.Text = A + Val(Text14.Text)
Text17.Text = Val(Text14.Text) + Val(Text16.Text)
End Sub

Private Sub Command3_Click()
Dim STR As String
STR = "INSERT INTO PURCHASE_MASTER VALUES('" & Text1.Text & "','" & Text3.Text & "','" & Text6.Text & "'," & Val(Text10.Text) & "," & Val(Text14.Text) & "," & Val(Text19.Text) & ")"
con.Execute STR
con.Execute "COMMIT"

For I = 0 To List6.ListCount - 1
con.Execute "INSERT INTO PURCHASE_ORD_DETAILS VALUES('" & Text1.Text & "','" & List2.List(I) & "'," & Val(List5.List(I)) & "," & Val(List6.List(I)) & ")"
con.Execute "COMMIT"
Next

Dim STR1 As String
STR1 = "UPDATE SUPPLIER SET S_DUES = " & Val(Text18.Text) & " WHERE S_NAME = '" & Combo3.Text & "' "
con.Execute STR1
con.Execute "COMMIT"

MsgBox "PURCHASE COMPLETED"
Unload Me
Load Me
End Sub

Private Sub Command4_Click()
Text14.Text = (Val(Text9.Text) - Val(List6.List(List6.ListCount - 1)))
List1.RemoveItem (List1.ListCount - 1)
List2.RemoveItem (List2.ListCount - 1)
List3.RemoveItem (List3.ListCount - 1)
List4.RemoveItem (List4.ListCount - 1)
List5.RemoveItem (List5.ListCount - 1)
List6.RemoveItem (List6.ListCount - 1)
Text10.Text = List1.ListCount - 1

End Sub

Private Sub Form_Activate()
C = 0

End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "select * from SUPPLIER ", con, adOpenDynamic
While rs.EOF = False
Combo3.AddItem rs.Fields(1)
rs.MoveNext
Wend
Text6.Text = Format(Date, "DD-MMM-YYYY")
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PURCHASE_MASTER", con, adOpenDynamic
If rs.BOF = True Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT count(PUR_NO) FROM PURCHASE_MASTER", con, adOpenDynamic
A = rs.Fields(0)
Text1.Text = "P" & "U" & "0" & (A + 1)
P = Text1.Text
ElseIf P >= "PU09" Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(PUR_NO) FROM PURCHASE_MASTER", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "P" & "U" & (A + 1)
P = Text1.Text
Else
If rs.State = 1 Then rs.Close
rs.Open "SELECT max(PUR_NO) FROM PURCHASE_MASTER", con, adOpenDynamic
A = Right(rs.Fields(0), 2)
Text1.Text = "P" & "U" & "0" & (A + 1)
P = Text1.Text
End If
End Sub



Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If rs.State = 1 Then rs.Close
rs.Open "select * from SUPPLIER_PRODUCT WHERE S_ID= '" & Text3.Text & "' AND P_NO ='" & Text8.Text & "' ", con, adOpenDynamic
Text9.Text = rs.Fields(2) * Val(Text11.Text)

End Sub

Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)
Text19.Text = Val(Text13.Text) - Val(Text16.Text)
Text18.Text = Val(Text14.Text) - Val(Text19.Text)
End Sub


Private Sub Text15_Change()

End Sub

