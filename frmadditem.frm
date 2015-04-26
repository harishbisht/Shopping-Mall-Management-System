VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   10695
   ClientLeft      =   1035
   ClientTop       =   450
   ClientWidth     =   13605
   LinkTopic       =   "Form9"
   ScaleHeight     =   10695
   ScaleWidth      =   13605
   Begin vkUserContolsXP.vkFrame frmproddetail 
      Height          =   1335
      Left            =   5880
      TabIndex        =   1
      Top             =   1320
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2355
      Caption         =   "Product Detail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   375
         Left            =   6000
         TabIndex        =   30
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Discount"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkLabel vkLabel6 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Type"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cmbtype 
         Height          =   315
         ItemData        =   "frmadditem.frx":0000
         Left            =   120
         List            =   "frmadditem.frx":0010
         TabIndex        =   12
         Text            =   "(Select)"
         Top             =   840
         Width           =   1575
      End
      Begin vkUserContolsXP.vkTextBox txtprodstock 
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkTextBox txtproddisc 
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         MaxLength       =   2
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkTextBox txtprodprice 
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkTextBox txtprodname 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkTextBox txtprodid 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkLabel vkLabel5 
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Stock"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Price"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Product Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Product ID"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid prodgrid 
      Height          =   4125
      Left            =   2040
      TabIndex        =   0
      Top             =   2760
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   7276
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1155
      Left            =   2040
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2037
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dist ID"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Dist Name"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Company Name"
         Object.Width           =   2452
      EndProperty
   End
   Begin vkUserContolsXP.vkFrame framefunctiom 
      Height          =   10455
      Left            =   240
      TabIndex        =   14
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   18441
      Caption         =   "Meubar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ButtonXp.XPButton cmdcancel 
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Cancel"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmddelete 
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Delete"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdedit 
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Edit"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdadd 
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Add"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdtop 
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   4440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Top"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdbottom 
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   6600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Bottom"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdnext 
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   5160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Next"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdprevious 
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   5880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Previous"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdclose 
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   9000
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Close"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdsearch 
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   8280
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Search"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdclosemenu 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   9960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Close Menu"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdsave 
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Save"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdendsearch 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   8280
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&End Search"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdclear 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   7560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C&lear"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   1320
         Y1              =   9840
         Y2              =   9840
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   1320
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   1320
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1320
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line Line6 
         X1              =   -240
         X2              =   1080
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   1320
         Y1              =   -360
         Y2              =   -360
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   1320
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   1320
         Y1              =   -240
         Y2              =   -240
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   1320
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line11 
         X1              =   0
         X2              =   1320
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line12 
         X1              =   0
         X2              =   1440
         Y1              =   9720
         Y2              =   9720
      End
   End
   Begin vkUserContolsXP.vkFrame framesearch 
      Height          =   2535
      Left            =   2520
      TabIndex        =   31
      Top             =   6960
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4471
      Caption         =   "Search Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   255
      TitleColor2     =   65535
      BorderColor     =   16711680
      Begin prjXTab.XTab XTab1 
         Height          =   1935
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3413
         TabCount        =   4
         TabCaption(0)   =   "Product ID"
         TabContCtrlCnt(0)=   3
         Tab(0)ContCtrlCap(1)=   "vkLabel12"
         Tab(0)ContCtrlCap(2)=   "cmdok1"
         Tab(0)ContCtrlCap(3)=   "txtsprodid"
         TabCaption(1)   =   "Product Name"
         TabContCtrlCnt(1)=   3
         Tab(1)ContCtrlCap(1)=   "cmdok2"
         Tab(1)ContCtrlCap(2)=   "txtsprodname"
         Tab(1)ContCtrlCap(3)=   "vkLabel9"
         TabCaption(2)   =   "Product Type"
         TabContCtrlCnt(2)=   3
         Tab(2)ContCtrlCap(1)=   "cmbstype"
         Tab(2)ContCtrlCap(2)=   "cmdok3"
         Tab(2)ContCtrlCap(3)=   "vkLabel10"
         TabCaption(3)   =   "Company Name"
         TabContCtrlCnt(3)=   3
         Tab(3)ContCtrlCap(1)=   "cmdok4"
         Tab(3)ContCtrlCap(2)=   "txtsprodcomp"
         Tab(3)ContCtrlCap(3)=   "vkLabel8"
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   -2147483627
         Begin VB.ComboBox cmbstype 
            Height          =   315
            ItemData        =   "frmadditem.frx":003E
            Left            =   -72240
            List            =   "frmadditem.frx":004E
            TabIndex        =   46
            Top             =   840
            Width           =   2175
         End
         Begin ButtonXp.XPButton cmdok4 
            Height          =   615
            Left            =   -69240
            TabIndex        =   45
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Command1"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin vkUserContolsXP.vkLabel vkLabel12 
            Height          =   375
            Left            =   1440
            TabIndex        =   44
            Top             =   840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Product ID"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
         Begin ButtonXp.XPButton cmdok3 
            Height          =   495
            Left            =   -69840
            TabIndex        =   43
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Command2"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin ButtonXp.XPButton cmdok2 
            Height          =   375
            Left            =   -69720
            TabIndex        =   42
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Command1"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin ButtonXp.XPButton cmdok1 
            Height          =   495
            Left            =   5160
            TabIndex        =   41
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Command1"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin vkUserContolsXP.vkLabel vkLabel11 
            Height          =   255
            Left            =   -73800
            TabIndex        =   40
            Top             =   480
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Distributor ID in the form of ""Dist1"""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin vkUserContolsXP.vkTextBox txtsprodname 
            Height          =   375
            Left            =   -71880
            TabIndex        =   39
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            LegendForeColor =   12937777
         End
         Begin vkUserContolsXP.vkLabel vkLabel9 
            Height          =   375
            Left            =   -74640
            TabIndex        =   38
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Product Name  "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
         Begin vkUserContolsXP.vkTextBox txtsprodid 
            Height          =   375
            Left            =   3720
            TabIndex        =   37
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Text            =   "P"
            BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            LegendForeColor =   12937777
         End
         Begin vkUserContolsXP.vkLabel vkLabel7 
            Height          =   375
            Left            =   -74040
            TabIndex        =   36
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Customer ID "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
         Begin vkUserContolsXP.vkTextBox txtsprodcomp 
            Height          =   375
            Left            =   -71400
            TabIndex        =   35
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            MaxLength       =   10
            LegendForeColor =   12937777
         End
         Begin vkUserContolsXP.vkLabel vkLabel8 
            Height          =   375
            Left            =   -74280
            TabIndex        =   34
            Top             =   960
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Company Name"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
         Begin vkUserContolsXP.vkLabel vkLabel10 
            Height          =   375
            Left            =   -74640
            TabIndex        =   33
            Top             =   840
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Product Type"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsprod As New ADODB.Recordset
Dim rsprod1 As New ADODB.Recordset
Dim rsprod2 As New ADODB.Recordset
Dim ff As Integer
Dim aa As Integer
Dim a, b, x, y As String
Dim sql As String
Dim sql1 As String
Dim sql2 As String
Dim addflag As Boolean

Private Sub cmbstype_Click()
Select Case cmbstype.ListIndex
Case 0:
x = "Type"
a = cmbstype.Text
Call search

Case 1:
x = "Type"
a = cmbstype.Text
Call search

Case 2:
x = "Type"
a = cmbstype.Text
Call search

Case 3:
x = "Type"
a = cmbstype.Text
Call search
End Select

End Sub


Private Sub Form_Load()
Call connection
sql = "select* from DistributorDetail"
rsprod.Open sql, cn, adOpenDynamic, adLockOptimistic

ff = 1
Do While Not rsprod.EOF
    ListView1.ListItems.Add , , rsprod.Fields("Distributor ID")
    ListView1.ListItems(ff).SubItems(1) = rsprod.Fields("Full Name")
    ListView1.ListItems(ff).SubItems(2) = rsprod.Fields("Company Name")
ff = ff + 1
rsprod.MoveNext
Loop
gridload
gridfill
rsprod.Close
End Sub

Private Sub gridload()

'''TO GENERATE PRODUCT GRID
prodgrid.TextMatrix(0, 0) = "No."
prodgrid.TextMatrix(0, 1) = "DistID"
prodgrid.TextMatrix(0, 2) = "Dist Name"
prodgrid.TextMatrix(0, 3) = "Company Name"
prodgrid.TextMatrix(0, 4) = "Type"
prodgrid.TextMatrix(0, 5) = "Prod ID"
prodgrid.TextMatrix(0, 6) = "Prod Name"
prodgrid.TextMatrix(0, 7) = "Price"
prodgrid.TextMatrix(0, 8) = "Stock"
prodgrid.TextMatrix(0, 9) = "Disc"

prodgrid.ColWidth(0) = 350
prodgrid.ColWidth(1) = 750
prodgrid.ColWidth(2) = 1200
prodgrid.ColWidth(3) = 1500
prodgrid.ColWidth(4) = 1800
prodgrid.ColWidth(5) = 1300
prodgrid.ColWidth(6) = 2100
prodgrid.ColWidth(7) = 800
prodgrid.ColWidth(8) = 700
prodgrid.ColWidth(9) = 575

End Sub

Private Sub gridfill()
'''FOR ADDING RECORDS FROM DATABASE TO GRID
sql1 = "select*from ProductDetail "
rsprod1.Open sql1, cn, adOpenDynamic, adLockOptimistic
prodgrid.Rows = 1
ff = 1
rsprod1.MoveFirst
Do While Not rsprod1.EOF
    prodgrid.Rows = prodgrid.Rows + 1
    prodgrid.TextMatrix(ff, 0) = ff
    prodgrid.TextMatrix(ff, 1) = IIf(IsNull(rsprod1.Fields("Dist ID")), "", rsprod1.Fields("Dist ID"))
'''-------
'''From List to grid
    rsprod.MoveFirst
    Do While Not rsprod.EOF
    If rsprod1.Fields("Dist ID") = rsprod.Fields("Distributor ID") Then
    prodgrid.TextMatrix(ff, 2) = IIf(IsNull(rsprod.Fields("Full Name")), "", rsprod.Fields("Full Name"))
    prodgrid.TextMatrix(ff, 3) = IIf(IsNull(rsprod.Fields("Company Name")), "", rsprod.Fields("Company Name"))
    Else
    End If
    rsprod.MoveNext
    Loop
'''-------
    prodgrid.TextMatrix(ff, 4) = IIf(IsNull(rsprod1.Fields("Type")), "", rsprod1.Fields("Type"))
    prodgrid.TextMatrix(ff, 5) = IIf(IsNull(rsprod1.Fields("Prod ID")), "", rsprod1.Fields("Prod ID"))
    prodgrid.TextMatrix(ff, 6) = IIf(IsNull(rsprod1.Fields("Prod Name")), "", rsprod1.Fields("Prod Name"))
    prodgrid.TextMatrix(ff, 7) = IIf(IsNull(rsprod1.Fields("Price")), "", rsprod1.Fields("Price"))
    prodgrid.TextMatrix(ff, 8) = IIf(IsNull(rsprod1.Fields("Stock")), "", rsprod1.Fields("Stock"))
    prodgrid.TextMatrix(ff, 9) = IIf(IsNull(rsprod1.Fields("Discount")), "", rsprod1.Fields("Discount"))
    rsprod1.MoveNext
ff = ff + 1
Loop

End Sub

Private Sub prodgrid_Click()
txtprodid.Text = prodgrid.TextMatrix(prodgrid.Row, 5)
txtprodname.Text = prodgrid.TextMatrix(prodgrid.Row, 6)
cmbtype.Text = prodgrid.TextMatrix(prodgrid.Row, 4)
txtprodprice.Text = prodgrid.TextMatrix(prodgrid.Row, 7)
txtproddisc.Text = prodgrid.TextMatrix(prodgrid.Row, 8)
txtprodstock.Text = prodgrid.TextMatrix(prodgrid.Row, 9)
End Sub

Private Sub ListView1_Click()
If framesearch.Visible = True Then
Exit Sub
Else
a = ListView1.SelectedItem.Text

Call selection
End If
'rsprod2.Close
End Sub

Private Sub selection()

prodgrid.Clear
gridload

sql2 = "select * from ProductDetail "
rsprod2.Open sql2, cn, adOpenDynamic, adLockOptimistic
rsprod2.Update
prodgrid.Rows = 1
ff = 1

rsprod2.MoveFirst
aa = 0
Do While Not rsprod2.EOF

If a = rsprod2.Fields("Dist ID") Then

sql = "select* from DistributorDetail"
rsprod.Open sql, cn, adOpenDynamic, adLockOptimistic

Do While Not rsprod.EOF
If rsprod.Fields("Distributor ID") = a Then
prodgrid.Rows = prodgrid.Rows + 1
aa = aa + 1
    prodgrid.TextMatrix(aa, 0) = aa
    prodgrid.TextMatrix(aa, 1) = IIf(IsNull(rsprod.Fields("Distributor Id")), "", rsprod.Fields("Distributor ID"))
    prodgrid.TextMatrix(aa, 2) = IIf(IsNull(rsprod.Fields("Full Name")), "", rsprod.Fields("Full Name"))
    prodgrid.TextMatrix(aa, 3) = IIf(IsNull(rsprod.Fields("Company Name")), "", rsprod.Fields("Company Name"))
    prodgrid.TextMatrix(aa, 4) = IIf(IsNull(rsprod2.Fields("Type")), "", rsprod2.Fields("Type"))
    prodgrid.TextMatrix(aa, 5) = IIf(IsNull(rsprod2.Fields("Prod ID")), "", rsprod2.Fields("Prod ID"))
    prodgrid.TextMatrix(aa, 6) = IIf(IsNull(rsprod2.Fields("Prod Name")), "", rsprod2.Fields("Prod Name"))
    prodgrid.TextMatrix(aa, 7) = IIf(IsNull(rsprod2.Fields("Price")), "", rsprod2.Fields("Price"))
    prodgrid.TextMatrix(aa, 8) = IIf(IsNull(rsprod2.Fields("Stock")), "", rsprod2.Fields("Stock"))
    prodgrid.TextMatrix(aa, 9) = IIf(IsNull(rsprod2.Fields("Discount")), "", rsprod2.Fields("Discount"))

rsprod.MoveNext
Else
rsprod.MoveNext
End If
Loop
rsprod.Close
rsprod2.MoveNext
Else
rsprod2.MoveNext
End If

Loop
rsprod2.Close
End Sub

Private Sub cmbtype_Click()
Select Case cmbtype.ListIndex
Case 0:
    b = "General"
    i = 1
Case 1:
    b = "Clothes"
    i = 2
Case 2:
    b = "Electronic"
    i = 3
Case 3:
    b = "Footware"
    i = 4
End Select
cmbsel

    If prodgrid.Rows = 1 Then
        txtprodid.Text = "P" & (i) & (prodgrid.Rows)
    Else
        txtprodid.Text = "P" & (i) & (prodgrid.Rows - 1 + 1)
        txtprodid.Enabled = False
    End If
    txtprodname.SetFocus
End Sub

Private Sub cmbsel()
prodgrid.Clear
gridload
sql2 = "select * from ProductDetail "
rsprod2.Open sql2, cn, adOpenDynamic, adLockOptimistic
rsprod2.Update
prodgrid.Rows = 1
ff = 1

rsprod2.MoveFirst
aa = 0
Do While Not rsprod2.EOF

If ListView1.SelectedItem.Text = rsprod2.Fields("Dist ID") Then
'rsprod.Close
sql = "select* from DistributorDetail"
rsprod.Open sql, cn, adOpenDynamic, adLockOptimistic

Do While Not rsprod.EOF
If rsprod.Fields("Distributor ID") = a And rsprod2.Fields("Type") = b Then
prodgrid.Rows = prodgrid.Rows + 1
aa = aa + 1
    prodgrid.TextMatrix(aa, 0) = aa
    prodgrid.TextMatrix(aa, 1) = IIf(IsNull(rsprod.Fields("Distributor Id")), "", rsprod.Fields("Distributor ID"))
    prodgrid.TextMatrix(aa, 2) = IIf(IsNull(rsprod.Fields("Full Name")), "", rsprod.Fields("Full Name"))
    prodgrid.TextMatrix(aa, 3) = IIf(IsNull(rsprod.Fields("Company Name")), "", rsprod.Fields("Company Name"))
    prodgrid.TextMatrix(aa, 4) = IIf(IsNull(rsprod2.Fields("Type")), "", rsprod2.Fields("Type"))
    prodgrid.TextMatrix(aa, 5) = IIf(IsNull(rsprod2.Fields("Prod ID")), "", rsprod2.Fields("Prod ID"))
    prodgrid.TextMatrix(aa, 6) = IIf(IsNull(rsprod2.Fields("Prod Name")), "", rsprod2.Fields("Prod Name"))
    prodgrid.TextMatrix(aa, 7) = IIf(IsNull(rsprod2.Fields("Price")), "", rsprod2.Fields("Price"))
    prodgrid.TextMatrix(aa, 8) = IIf(IsNull(rsprod2.Fields("Stock")), "", rsprod2.Fields("Stock"))
    prodgrid.TextMatrix(aa, 9) = IIf(IsNull(rsprod2.Fields("Discount")), "", rsprod2.Fields("Discount"))
    
rsprod.MoveNext
Else
rsprod.MoveNext
End If
Loop

rsprod2.MoveNext
rsprod.Close
Else
rsprod2.MoveNext
'rsprod.Close
End If
Loop
rsprod2.Close
End Sub

Private Sub cmdsave_Click()
cmdsave.Enabled = False
addflag = True
flagopen = False
If flagopen = True Then
'rsprod1.Close

End If

If addflag = True Then
        rsprod1.AddNew
        rsprod1.Fields("Dist ID") = ListView1.SelectedItem.Text
        rsprod1.Fields("Type") = cmbtype.Text
        rsprod1.Fields("Prod ID") = txtprodid.Text
        rsprod1.Fields("Prod Name") = txtprodname.Text
        rsprod1.Fields("Price") = txtprodprice.Text
        rsprod1.Fields("Discount") = txtproddisc.Text
        rsprod1.Fields("Stock") = txtprodstock.Text
        rsprod1.Update
End If
Call gridsave
rsprod1.Update
'rsprod1.Close
flagopen = False

End Sub

Private Sub gridsave()
prodgrid.Rows = prodgrid.Rows + 1
If prodgrid.Rows = 2 Then
prodgrid.TextMatrix(prodgrid.Rows - 1, 0) = 1
Else
prodgrid.TextMatrix(prodgrid.Rows - 1, 0) = prodgrid.TextMatrix(prodgrid.Rows - 2, 0) + 1
End If
    prodgrid.TextMatrix(prodgrid.Rows - 1, 1) = ListView1.SelectedItem.Text
    prodgrid.TextMatrix(prodgrid.Rows - 1, 2) = ListView1.SelectedItem.SubItems(1)
    prodgrid.TextMatrix(prodgrid.Rows - 1, 3) = ListView1.SelectedItem.SubItems(2)
    prodgrid.TextMatrix(prodgrid.Rows - 1, 4) = cmbtype.Text
    prodgrid.TextMatrix(prodgrid.Rows - 1, 5) = txtprodid.Text
    prodgrid.TextMatrix(prodgrid.Rows - 1, 6) = txtprodname.Text
    prodgrid.TextMatrix(prodgrid.Rows - 1, 7) = txtprodprice.Text
    prodgrid.TextMatrix(prodgrid.Rows - 1, 8) = txtproddisc.Text
    prodgrid.TextMatrix(prodgrid.Rows - 1, 9) = txtprodstock.Text
'    prodgrid.Refresh
End Sub

Private Sub cmdtop_Click()
    prodgrid.Row = 1
    prodgrid.SetFocus
  prodgrid_Click
End Sub

Private Sub cmdnext_Click()
If prodgrid.Row < prodgrid.Rows - 1 Then
   prodgrid.Row = prodgrid.Row + 1
   prodgrid_Click
   cmdnext.SetFocus
End If
End Sub

Private Sub cmdprevious_Click()
If prodgrid.Row > 1 Then
    If prodgrid.Row - 1 < prodgrid.Rows Then
        prodgrid.Row = prodgrid.Row - 1
    prodgrid_Click
        cmdprevious.SetFocus
    End If
End If
End Sub

Private Sub cmdbottom_Click()
prodgrid.Row = prodgrid.Rows - 1
prodgrid_Click
End Sub

Private Sub cmdclear_Click()
cmbtype.Text = ""
txtprodid.Text = ""
txtprodname.Text = ""
txtprodprice.Text = ""
txtprodstock.Text = ""
txtproddisc.Text = ""
txtsprodid.Text = "P"
txtsprodname.Text = ""
cmbstype.Text = ""
txtsprodcomp.Text = ""

End Sub

Private Sub cmdsearch_Click()
framesearch.Visible = True
cmdsearch.Visible = False
cmdendsearch.Visible = True
End Sub

Private Sub cmdendsearch_Click()
framesearch.Visible = False
cmdsearch.Visible = True
cmdendsearch.Visible = False
End Sub

Private Sub cmdok1_Click()
x = "Prod ID"
a = txtsprodid.Text
Call search
End Sub

Private Sub cmdok2_Click()
x = "Prod Name"
a = txtsprodname.Text
Call search
End Sub

Private Sub search()
prodgrid.Clear
gridload
'rsprod2.Close
sql2 = "select * from ProductDetail "
rsprod2.Open sql2, cn, adOpenDynamic, adLockOptimistic
rsprod2.Update
prodgrid.Rows = 1
ff = 1

rsprod2.MoveFirst
aa = 0

Do While Not rsprod2.EOF
    If rsprod2.Fields(x) = a Then
    
        prodgrid.Rows = prodgrid.Rows + 1
        aa = aa + 1
        
            prodgrid.TextMatrix(aa, 0) = aa
            prodgrid.TextMatrix(aa, 1) = IIf(IsNull(rsprod2.Fields("Dist ID")), "", rsprod2.Fields("Dist ID"))
        
        sql = "select * from DistributorDetail "
        
        rsprod.Open sql, cn, adOpenDynamic, adLockOptimistic
        rsprod.MoveFirst
        Do While Not rsprod.EOF
            If rsprod.Fields("Distributor ID") = rsprod2.Fields("Dist ID") Then
            
                prodgrid.TextMatrix(aa, 2) = IIf(IsNull(rsprod.Fields("Full Name")), "", rsprod.Fields("Full Name"))
                prodgrid.TextMatrix(aa, 3) = IIf(IsNull(rsprod.Fields("Company Name")), "", rsprod.Fields("Company Name"))
                rsprod.MoveNext
            Else
                rsprod.MoveNext
            End If
        Loop
            rsprod.Close
            prodgrid.TextMatrix(aa, 4) = IIf(IsNull(rsprod2.Fields("Type")), "", rsprod2.Fields("Type"))
            prodgrid.TextMatrix(aa, 5) = IIf(IsNull(rsprod2.Fields("Prod ID")), "", rsprod2.Fields("Prod ID"))
            prodgrid.TextMatrix(aa, 6) = IIf(IsNull(rsprod2.Fields("Prod Name")), "", rsprod2.Fields("Prod Name"))
            prodgrid.TextMatrix(aa, 7) = IIf(IsNull(rsprod2.Fields("Price")), "", rsprod2.Fields("Price"))
            prodgrid.TextMatrix(aa, 8) = IIf(IsNull(rsprod2.Fields("Stock")), "", rsprod2.Fields("Stock"))
            prodgrid.TextMatrix(aa, 9) = IIf(IsNull(rsprod2.Fields("Discount")), "", rsprod2.Fields("Discount"))
        
        rsprod2.MoveNext
    Else
        rsprod2.MoveNext
    End If
    If rsprod2.EOF = True Then
        rsprod2.Close
        Exit Sub
    Else
    End If
Loop
rsprod.Close
End Sub

Private Sub cmdok4_Click()

prodgrid.Clear
gridload
sql = "select* from DistributorDetail"
rsprod.Open sql, cn, adOpenDynamic, adLockOptimistic
prodgrid.Rows = 1
ff = 1
rsprod.MoveFirst
aa = 0
 sql2 = "select * from ProductDetail "
        rsprod2.Open sql2, cn, adOpenDynamic, adLockOptimistic
        rsprod2.MoveFirst
Do While rsprod.EOF = False
If rsprod.Fields("Company Name") = txtsprodcomp.Text Then
        prodgrid.Rows = prodgrid.Rows + 1
        aa = aa + 1
        
            prodgrid.TextMatrix(aa, 0) = aa
            prodgrid.TextMatrix(aa, 1) = IIf(IsNull(rsprod.Fields("Distributor ID")), "", rsprod.Fields("Distributor ID"))
            prodgrid.TextMatrix(aa, 2) = IIf(IsNull(rsprod.Fields("Full Name")), "", rsprod.Fields("Full Name"))
            prodgrid.TextMatrix(aa, 3) = IIf(IsNull(rsprod.Fields("Company Name")), "", rsprod.Fields("Company Name"))

rsprod2.MoveFirst

        Do While Not rsprod2.EOF
            If rsprod2.Fields("Dist ID") = rsprod.Fields("Distributor ID") Then
            
            prodgrid.TextMatrix(aa, 4) = IIf(IsNull(rsprod2.Fields("Type")), "", rsprod2.Fields("Type"))
            prodgrid.TextMatrix(aa, 5) = IIf(IsNull(rsprod2.Fields("Prod ID")), "", rsprod2.Fields("Prod ID"))
            prodgrid.TextMatrix(aa, 6) = IIf(IsNull(rsprod2.Fields("Prod Name")), "", rsprod2.Fields("Prod Name"))
            prodgrid.TextMatrix(aa, 7) = IIf(IsNull(rsprod2.Fields("Price")), "", rsprod2.Fields("Price"))
            prodgrid.TextMatrix(aa, 8) = IIf(IsNull(rsprod2.Fields("Stock")), "", rsprod2.Fields("Stock"))
            prodgrid.TextMatrix(aa, 9) = IIf(IsNull(rsprod2.Fields("Discount")), "", rsprod2.Fields("Discount"))
            
                rsprod2.MoveNext
            Else
                rsprod2.MoveNext
            End If
        Loop
        rsprod.MoveNext
    Else
        rsprod.MoveNext
    End If
    If rsprod.EOF = True Then
        rsprod.Close
        rsprod2.Close
        Exit Sub
    Else
    
    End If
    Loop
rsprod2.Close
End Sub


Private Sub txtprodprice_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 8 Or KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtproddisc_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 8 Or KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtprodstock_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 8 Or KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub
