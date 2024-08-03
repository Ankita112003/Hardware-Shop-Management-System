VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form15 
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   22920
   LinkTopic       =   "Form15"
   ScaleHeight     =   12495
   ScaleWidth      =   22920
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
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
      Left            =   18840
      TabIndex        =   66
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
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
      Left            =   14040
      TabIndex        =   65
      Top             =   10680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
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
      Left            =   14040
      TabIndex        =   64
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
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
      Left            =   18840
      TabIndex        =   63
      Top             =   9240
      Width           =   2055
   End
   Begin VB.ListBox List11 
      Height          =   2595
      ItemData        =   "Form15.frx":0000
      Left            =   11640
      List            =   "Form15.frx":0002
      TabIndex        =   61
      Top             =   9480
      Width           =   1695
   End
   Begin VB.ListBox List9 
      Height          =   2595
      ItemData        =   "Form15.frx":0004
      Left            =   9960
      List            =   "Form15.frx":0006
      TabIndex        =   60
      Top             =   9480
      Width           =   1695
   End
   Begin VB.ListBox List5 
      Height          =   2595
      ItemData        =   "Form15.frx":0008
      Left            =   8040
      List            =   "Form15.frx":000A
      TabIndex        =   59
      Top             =   9480
      Width           =   1935
   End
   Begin VB.ListBox List4 
      Height          =   2595
      ItemData        =   "Form15.frx":000C
      Left            =   6120
      List            =   "Form15.frx":000E
      TabIndex        =   58
      Top             =   9480
      Width           =   1935
   End
   Begin VB.ListBox List3 
      Height          =   2595
      ItemData        =   "Form15.frx":0010
      Left            =   4320
      List            =   "Form15.frx":0012
      TabIndex        =   57
      Top             =   9480
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   2595
      ItemData        =   "Form15.frx":0014
      Left            =   3480
      List            =   "Form15.frx":0016
      TabIndex        =   56
      Top             =   9480
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      TabIndex        =   49
      Text            =   "Choose product no."
      Top             =   8280
      Width           =   2055
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   14400
      TabIndex        =   48
      Top             =   8160
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   495
      Left            =   12000
      TabIndex        =   47
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   495
      Left            =   9480
      TabIndex        =   46
      Top             =   8160
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      TabIndex        =   45
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   44
      Top             =   8160
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   4080
      TabIndex        =   37
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      Format          =   205455361
      CurrentDate     =   45352
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      TabIndex        =   36
      Text            =   "CHOOSE PURCHASE NO."
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox Text15 
      Height          =   615
      Left            =   19800
      TabIndex        =   35
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   615
      Left            =   17040
      TabIndex        =   34
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      Height          =   615
      Left            =   19800
      TabIndex        =   33
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   615
      Left            =   17040
      TabIndex        =   32
      Top             =   4080
      Width           =   1815
   End
   Begin VB.ListBox List12 
      Height          =   2400
      ItemData        =   "Form15.frx":0018
      Left            =   12840
      List            =   "Form15.frx":001A
      TabIndex        =   26
      Top             =   4200
      Width           =   1695
   End
   Begin VB.ListBox List10 
      Height          =   2400
      ItemData        =   "Form15.frx":001C
      Left            =   11520
      List            =   "Form15.frx":001E
      TabIndex        =   25
      Top             =   4200
      Width           =   1335
   End
   Begin VB.ListBox List8 
      Height          =   2400
      ItemData        =   "Form15.frx":0020
      Left            =   10080
      List            =   "Form15.frx":0022
      TabIndex        =   24
      Top             =   4200
      Width           =   1455
   End
   Begin VB.ListBox List7 
      Height          =   2400
      ItemData        =   "Form15.frx":0024
      Left            =   8040
      List            =   "Form15.frx":0026
      TabIndex        =   23
      Top             =   4200
      Width           =   2055
   End
   Begin VB.ListBox List6 
      Height          =   2400
      ItemData        =   "Form15.frx":0028
      Left            =   6120
      List            =   "Form15.frx":002A
      TabIndex        =   22
      Top             =   4200
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "Form15.frx":002C
      Left            =   5400
      List            =   "Form15.frx":002E
      TabIndex        =   21
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   615
      Left            =   17640
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   615
      Left            =   17640
      TabIndex        =   12
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   615
      Left            =   13080
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   615
      Left            =   12600
      TabIndex        =   10
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   8880
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   615
      Left            =   8880
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.Shape Shape10 
      BorderWidth     =   3
      Height          =   2055
      Left            =   18600
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Shape Shape9 
      Height          =   855
      Left            =   18720
      Top             =   9120
      Width           =   2295
   End
   Begin VB.Shape Shape8 
      Height          =   855
      Left            =   18720
      Top             =   8160
      Width           =   2295
   End
   Begin VB.Shape Shape7 
      BorderWidth     =   3
      Height          =   2175
      Left            =   13800
      Top             =   9360
      Width           =   2535
   End
   Begin VB.Shape Shape6 
      Height          =   855
      Left            =   13920
      Top             =   10560
      Width           =   2295
   End
   Begin VB.Shape Shape5 
      Height          =   855
      Left            =   13920
      Top             =   9480
      Width           =   2295
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Caption         =   "PRODUCT RECEIVED"
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
      Left            =   2640
      TabIndex        =   62
      Top             =   6960
      Width           =   4455
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   3
      Height          =   5175
      Left            =   2040
      Top             =   7080
      Width           =   14655
   End
   Begin VB.Label Label32 
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
      Height          =   495
      Left            =   11640
      TabIndex        =   55
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label Label31 
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
      Height          =   495
      Left            =   9960
      TabIndex        =   54
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label Label30 
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
      Height          =   495
      Left            =   8040
      TabIndex        =   53
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label Label29 
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
      Height          =   495
      Left            =   6120
      TabIndex        =   52
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   4320
      TabIndex        =   51
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   3480
      TabIndex        =   50
      Top             =   9000
      Width           =   855
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY RECEIVED"
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
      Left            =   14160
      TabIndex        =   43
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY PURCHASED"
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
      Left            =   11640
      TabIndex        =   42
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE RATE"
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
      Left            =   9480
      TabIndex        =   41
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label23 
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
      Left            =   7440
      TabIndex        =   40
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label22 
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
      Left            =   5040
      TabIndex        =   39
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT_NO."
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
      Left            =   2640
      TabIndex        =   38
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   3495
      Left            =   16440
      Top             =   3120
      Width           =   5535
   End
   Begin VB.Label Label20 
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
      Height          =   615
      Left            =   19560
      TabIndex        =   31
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label19 
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
      Height          =   615
      Left            =   16920
      TabIndex        =   30
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADVANCE PAYMENT"
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
      Left            =   19680
      TabIndex        =   29
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label17 
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
      Height          =   615
      Left            =   17040
      TabIndex        =   28
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "PRODUCT PURCHASED"
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
      Left            =   4800
      TabIndex        =   27
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   3735
      Left            =   4920
      Top             =   3120
      Width           =   9855
   End
   Begin VB.Label Label15 
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
      Height          =   615
      Left            =   12840
      TabIndex        =   20
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label14 
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
      Height          =   615
      Left            =   11520
      TabIndex        =   19
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label13 
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
      Height          =   615
      Left            =   10080
      TabIndex        =   18
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label12 
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
      Height          =   615
      Left            =   8040
      TabIndex        =   17
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   6120
      TabIndex        =   16
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   5400
      TabIndex        =   15
      Top             =   3600
      Width           =   735
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
      Left            =   2400
      TabIndex        =   14
      Top             =   240
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   2295
      Left            =   1920
      Top             =   480
      Width           =   17655
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIVED DATE"
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
      Left            =   15360
      TabIndex        =   7
      Top             =   1920
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
      Left            =   15480
      TabIndex        =   6
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label6 
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
      Left            =   11160
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
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
      Left            =   11040
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   6360
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER ID"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE_NO"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE DATE"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AA As Double
Dim E As Integer

Private Sub Combo1_Click()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PURCHASE_MASTER WHERE PUR_NO ='" & Combo1.Text & "' ", con, adOpenDynamic
Text3.Text = rs.Fields(1)
Text13.Text = rs.Fields(5)
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM SUPPLIER WHERE S_ID ='" & Text3.Text & "' ", con, adOpenDynamic
Text1.Text = rs.Fields(1)
Text4.Text = rs.Fields(4)
Text5.Text = rs.Fields(2)
Text2.Text = rs.Fields(5)
For I = 0 To List1.ListCount - 1
List1.List(I) = ""
Next
If rs.State = 1 Then rs.Close
rs.Open "SELECT COUNT(*) FROM PURCHASE_ORD_DETAILS WHERE PUR_NO ='" & Combo1.Text & "' ", con, adOpenDynamic
C = rs.Fields(0)
For I = 0 To C - 1
List1.List(I) = (I + 1)
Next
For I = 0 To List6.ListCount - 1
List6.List(I) = ""
List10.List(I) = ""
List12.List(I) = ""
List7.List(I) = ""
List8.List(I) = ""
Next
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PURCHASE_ORD_DETAILS WHERE PUR_NO ='" & Combo1.Text & "' ", con, adOpenDynamic
For I = 0 To C - 1
List6.List(I) = rs.Fields(1)
List10.List(I) = rs.Fields(2)
List12.List(I) = rs.Fields(3)
rs.MoveNext
Next
For I = 0 To C - 1
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM Product WHERE P_no ='" & List6.List(I) & "' ", con, adOpenDynamic
List7.List(I) = rs.Fields(1)
List8.List(I) = rs.Fields(2)
Next
Combo2.Clear
Combo2.Text = "CHOOSE PRODUCT_NO"
If rs.State = 1 Then rs.Close
rs.Open "SELECT P_NO FROM PURCHASE_ORD_DETAILS WHERE PUR_NO ='" & Combo1.Text & "' ", con, adOpenDynamic
While rs.EOF = False
Combo2.AddItem rs.Fields(0)
rs.MoveNext
Wend
End Sub

Private Sub Combo2_Click()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM Product WHERE P_no ='" & Combo2.Text & "' ", con, adOpenDynamic
Text7.Text = rs.Fields(1)
Text8.Text = rs.Fields(2)
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM SUPPLIER_Product WHERE P_no ='" & Combo2.Text & "' AND S_ID = '" & Text3.Text & "' ", con, adOpenDynamic
Text9.Text = rs.Fields(2)
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PURCHASE_ORD_DETAILS WHERE PUR_NO ='" & Combo1.Text & "' ", con, adOpenDynamic
Text10.Text = rs.Fields(2)
Combo1.Enabled = False
End Sub

Private Sub Command1_Click()


Dim STR As String
STR = "INSERT INTO RECEIVED_MASTER VALUES('" & Combo1.Text & "','" & Text3.Text & "','" & Text6.Text & "'," & Val(List3.ListCount) & "," & Val(Text12.Text) & "," & (Val(Text12.Text) + Val(Text15.Text)) & ")"
con.Execute STR
con.Execute "COMMIT"


For I = 0 To List3.ListCount - 1
con.Execute "INSERT INTO RECEIVED_ORD_DETAILS VALUES('" & Combo1.Text & "','" & List3.List(I) & "'," & Val(List9.List(I)) & "," & Val(List11.List(I)) & ")"
con.Execute "COMMIT"
Next


Dim STR1 As String
STR1 = "UPDATE SUPPLIER SET S_DUES = " & Val(Text14.Text) & " WHERE S_NAME = '" & Text1.Text & "' "
con.Execute STR1
con.Execute "COMMIT"


Dim STR2 As String
For I = 0 To List3.ListCount - 1
STR2 = "UPDATE STOCK SET AVL_QUANTITY = " & Val(List9.List(I)) & " WHERE P_NO = '" & List3.List(I) & "' "
con.Execute STR2
con.Execute "COMMIT"
Next

MsgBox "SAVED"
Unload Me
Load Me
End Sub

Private Sub Command2_Click()
E = E + 1
List2.AddItem (E)
List3.AddItem Combo2.Text
List4.AddItem Text7.Text
List5.AddItem Text8.Text
List9.AddItem Text11.Text
List11.AddItem (Val(Text9.Text) * Val(Text11.Text))
K = Val(Text9.Text) * Val(Text11.Text)
Text12.Text = K + Val(Text12.Text)
Text14.Text = Val(Text12.Text) - Val(Text13.Text)
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Combo2.Text = "CHOOSE PRODUCT_NO"
AA = Val(Text14.Text)
End Sub



Private Sub Command3_Click()
List1.RemoveItem (List1.ListCount - 1)
List2.RemoveItem (List2.ListCount - 1)
List3.RemoveItem (List3.ListCount - 1)
List4.RemoveItem (List4.ListCount - 1)
List5.RemoveItem (List5.ListCount - 1)
List9.RemoveItem (List9.ListCount - 1)
List11.RemoveItem (List11.ListCount - 1)
End Sub

Private Sub Command4_Click()
Unload Me
Load Me
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "select * from PURCHASE_MASTER ", con, adOpenDynamic
While rs.EOF = False
Combo1.AddItem rs.Fields(0)
rs.MoveNext
Wend
Text6.Text = Format(Date, "DD-MMM-YYYY")
End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
Text14.Text = AA - Val(Text15.Text)
End Sub


