VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Distributer Details"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form6"
   ScaleHeight     =   9405
   ScaleWidth      =   14835
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Left            =   13080
      Top             =   9960
   End
   Begin VB.Timer Timer1 
      Left            =   12360
      Top             =   9960
   End
   Begin vkUserContolsXP.vkFrame frmdistdetail 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   3836
      Caption         =   "Distributer  Details"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkTextBox txtdistcaddress 
         Height          =   855
         Left            =   8280
         TabIndex        =   12
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1508
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   375
         Left            =   8280
         TabIndex        =   11
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Company Address"
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
         Left            =   6360
         TabIndex        =   10
         Top             =   600
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Company Name"
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
      Begin vkUserContolsXP.vkTextBox txtdistcname 
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Top             =   960
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
      Begin vkUserContolsXP.vkLabel Cust 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Distributor ID"
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
      Begin vkUserContolsXP.vkTextBox txtdistaddress 
         Height          =   855
         Left            =   3240
         TabIndex        =   7
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1508
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkTextBox txtdistid 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
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
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   600
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Address"
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
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Name"
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
      Begin vkUserContolsXP.vkTextBox txtdistname 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
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
      Begin vkUserContolsXP.vkTextBox txtdistmobileno 
         Height          =   375
         Left            =   11400
         TabIndex        =   2
         Top             =   960
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
         MaxLength       =   10
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkLabel vkLabel5 
         Height          =   375
         Left            =   11400
         TabIndex        =   1
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Mobile Number"
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
   Begin vkUserContolsXP.vkFrame framefunctiom 
      Height          =   10365
      Left            =   13800
      TabIndex        =   13
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   18283
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   9840
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   28
         Top             =   0
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   1320
         Y1              =   9720
         Y2              =   9720
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
         Y1              =   9600
         Y2              =   9600
      End
   End
   Begin MSFlexGridLib.MSFlexGrid distgrid 
      Height          =   2775
      Left            =   240
      TabIndex        =   29
      Top             =   2520
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin vkUserContolsXP.vkFrame framesearch 
      Height          =   2535
      Left            =   2400
      TabIndex        =   30
      Top             =   5520
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
      Begin prjXTab.XTab XTab1 
         Height          =   1935
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3413
         TabCount        =   4
         TabCaption(0)   =   "Distributor ID"
         TabContCtrlCnt(0)=   4
         Tab(0)ContCtrlCap(1)=   "vkLabel11"
         Tab(0)ContCtrlCap(2)=   "txtsdistid"
         Tab(0)ContCtrlCap(3)=   "vkLabel9"
         Tab(0)ContCtrlCap(4)=   "cmdok"
         TabCaption(1)   =   "Distributor Name"
         TabContCtrlCnt(1)=   3
         Tab(1)ContCtrlCap(1)=   "vkLabel7"
         Tab(1)ContCtrlCap(2)=   "txtsdistname"
         Tab(1)ContCtrlCap(3)=   "XPButton1"
         TabCaption(2)   =   "Company Name"
         TabContCtrlCnt(2)=   3
         Tab(2)ContCtrlCap(1)=   "vkLabel10"
         Tab(2)ContCtrlCap(2)=   "txtsdistcname"
         Tab(2)ContCtrlCap(3)=   "XPButton2"
         TabCaption(3)   =   "Mobile Number"
         TabContCtrlCnt(3)=   2
         Tab(3)ContCtrlCap(1)=   "vkLabel8"
         Tab(3)ContCtrlCap(2)=   "txtsdistmobile"
         ActiveTab       =   3
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
         Begin vkUserContolsXP.vkLabel vkLabel11 
            Height          =   255
            Left            =   -73680
            TabIndex        =   44
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Distributor ID"
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
         Begin vkUserContolsXP.vkLabel vkLabel10 
            Height          =   375
            Left            =   -74520
            TabIndex        =   43
            Top             =   720
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Company Name  "
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
         Begin vkUserContolsXP.vkTextBox txtsdistcname 
            Height          =   375
            Left            =   -72120
            TabIndex        =   42
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
         Begin vkUserContolsXP.vkLabel vkLabel8 
            Height          =   375
            Left            =   600
            TabIndex        =   41
            Top             =   960
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Mobile Number"
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
         Begin vkUserContolsXP.vkTextBox txtsdistmobile 
            Height          =   375
            Left            =   2880
            TabIndex        =   40
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
         Begin vkUserContolsXP.vkLabel vkLabel6 
            Height          =   375
            Left            =   -74040
            TabIndex        =   39
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
         Begin vkUserContolsXP.vkTextBox txtsdistid 
            Height          =   375
            Left            =   -71520
            TabIndex        =   38
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Text            =   "Dist"
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
            Left            =   -74640
            TabIndex        =   37
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Distributor Name  "
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
         Begin vkUserContolsXP.vkTextBox txtsdistname 
            Height          =   375
            Left            =   -71880
            TabIndex        =   36
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
            Height          =   255
            Left            =   -73800
            TabIndex        =   35
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
         Begin ButtonXp.XPButton cmdok 
            Height          =   255
            Left            =   -69240
            TabIndex        =   34
            Top             =   1080
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
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
         Begin ButtonXp.XPButton XPButton1 
            Height          =   375
            Left            =   -72480
            TabIndex        =   33
            Top             =   1320
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
         Begin ButtonXp.XPButton XPButton2 
            Height          =   615
            Left            =   -72960
            TabIndex        =   32
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "Command2"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim i As Integer
Dim ID As Integer
Dim ff As Integer
Dim sql As String
Dim j As Integer
Dim x As String
Dim y As String
Dim str As String
Dim addflag As Boolean
Dim click

Private Sub cmdadd_Click()
frmdistdetail.Enabled = True
cmdclear_Click
cmdbottom.Enabled = False
cmdtop.Enabled = False
cmdprevious.Enabled = False
cmdnext.Enabled = False
cmdedit.Enabled = False
cmdsave.Enabled = True
cmdsearch.Enabled = False
cmddelete.Enabled = False
cmdendsearch.Enabled = False
cmdadd.Enabled = False
addflag = True

'''TO GENERATE AUTOMATIC DISTRIBUTOR ID
    If distgrid.Rows = 1 Then
        txtdistid.Text = "dist1"
'    ElseIf distgrid.Rows = 2 Then
'        txtdistid.Text = "dist2"
    Else
        txtdistid.Text = "Dist" & (distgrid.Rows - 1 + 1)
        txtdistid.Enabled = False
    End If
    txtdistname.SetFocus
End Sub

Private Sub cmddelete_Click()
If distgrid.Rows = 1 Then
MsgBox "No Record found in Database"
Exit Sub
Else
click = MsgBox("Do u want to Delete this Record from database?", vbYesNo + vbQuestion)
If click = vbYes Then
rs.MoveFirst
While rs.EOF = False
If rs.Fields("Distributor ID") = txtdistid.Text Then
rs.Delete

distgrid.Clear
'distgrid.Rows = 1
'ff = 1
rs.MoveFirst
gridload
cmdclear_Click
Exit Sub
Else
rs.MoveNext
End If
Wend
End If
End If
End Sub

Private Sub cmdedit_Click()
frmdistdetail.Enabled = True
addflag = False
cmdsave.Enabled = True
txtdistid.Enabled = False
End Sub

Private Sub cmdCancel_Click()
cmdclear_Click
If cmdadd.Enabled = False Then
cmdbottom.Enabled = True
cmdtop.Enabled = True
cmdprevious.Enabled = True
cmdnext.Enabled = True
cmdedit.Enabled = True
cmdsave.Enabled = True
cmdsearch.Enabled = True
cmddelete.Enabled = True
cmdendsearch.Enabled = True
cmdadd.Enabled = True
End If
End Sub

Private Sub cmdsave_Click()
If txtdistname.Text = "" Or txtdistaddress.Text = "" Or txtdistcaddress.Text = "" Or txtdistcname.Text = "" Or txtdistmobileno.Text = "" Then
MsgBox "Please fill all the information.", vbInformation, "Field missing"
Exit Sub
End If

cmdsave.Enabled = False
cmdbottom.Enabled = True
cmdtop.Enabled = True
cmdprevious.Enabled = True
cmdnext.Enabled = True
cmdedit.Enabled = True
cmdsearch.Enabled = True
cmddelete.Enabled = True
cmdendsearch.Enabled = True
cmdadd.Enabled = True

''' FOR ADDING RECORD TO DATABASE AND GRID
If addflag = True Then
        rs.AddNew
        rs.Fields("distributor ID") = txtdistid.Text
        rs.Fields("Full Name") = txtdistname.Text
        rs.Fields("Address") = txtdistaddress.Text
        rs.Fields("Company Name") = txtdistcname.Text
        rs.Fields("Company Address") = txtdistcaddress.Text
        rs.Fields("Mobile") = txtdistmobileno.Text
        rs.Update
    Call gridfill

Else
''''' FOR EDITING RECORDS IN DATABASE AND GRID
        rs.MoveFirst
        Do While Not rs.EOF
         If rs.Fields("Distributor ID") = txtdistid.Text Then
        Exit Do
         End If
        rs.MoveNext
        Loop
        
''' TO EDIT DATABASE
        rs.Fields("Distributor ID") = txtdistid.Text
        rs.Fields("Full Name") = txtdistname.Text
        rs.Fields("Address") = txtdistaddress.Text
        rs.Fields("Company Name") = txtdistcname.Text
        rs.Fields("Company Address") = txtdistcaddress.Text
        rs.Fields("Mobile") = txtdistmobileno.Text
        rs.Update
        
''' TO FILL GRID FROM DATABASE
    distgrid.TextMatrix(distgrid.Row, 1) = rs.Fields("Distributor ID")
    distgrid.TextMatrix(distgrid.Row, 2) = rs.Fields("Full Name")
    distgrid.TextMatrix(distgrid.Row, 3) = rs.Fields("Address")
    distgrid.TextMatrix(distgrid.Row, 4) = rs.Fields("Company Name")
    distgrid.TextMatrix(distgrid.Row, 5) = rs.Fields("Company Address")
    distgrid.TextMatrix(distgrid.Row, 6) = rs.Fields("Mobile")
End If
cmdclear_Click
End Sub

Private Sub gridfill()
distgrid.Rows = distgrid.Rows + 1
If distgrid.Rows = 2 Then
distgrid.TextMatrix(distgrid.Rows - 1, 0) = 1
Else
distgrid.TextMatrix(distgrid.Rows - 1, 0) = distgrid.TextMatrix(distgrid.Rows - 2, 0) + 1
End If
    
    distgrid.TextMatrix(distgrid.Rows - 1, 1) = txtdistid.Text
    distgrid.TextMatrix(distgrid.Rows - 1, 2) = txtdistname.Text
    distgrid.TextMatrix(distgrid.Rows - 1, 3) = txtdistaddress.Text
    distgrid.TextMatrix(distgrid.Rows - 1, 4) = txtdistcname.Text
    distgrid.TextMatrix(distgrid.Rows - 1, 5) = txtdistcaddress.Text
    distgrid.TextMatrix(distgrid.Rows - 1, 6) = txtdistmobileno.Text
End Sub

Private Sub cmdclear_Click()
txtdistid.Text = ""
txtdistname.Text = ""
txtdistaddress.Text = ""
txtdistcname.Text = ""
txtdistcaddress.Text = ""
txtdistmobileno.Text = ""
End Sub

Private Sub cmdtop_Click()
    distgrid.Row = 1
    distgrid.SetFocus
  distgrid_Click
End Sub

Private Sub cmdnext_Click()
If distgrid.Row < distgrid.Rows - 1 Then
   distgrid.Row = distgrid.Row + 1
   distgrid_Click
   cmdnext.SetFocus
End If
End Sub

Private Sub cmdprevious_Click()
If distgrid.Row > 1 Then
    If distgrid.Row - 1 < distgrid.Rows Then
        distgrid.Row = distgrid.Row - 1
    distgrid_Click
        cmdprevious.SetFocus
    End If
End If
End Sub

Private Sub cmdbottom_Click()
distgrid.Row = distgrid.Rows - 1
distgrid_Click
End Sub

Private Sub cmdsearch_Click()
framesearch.Visible = True
cmdsearch.Visible = False
cmdendsearch.Visible = True
End Sub

Private Sub cmdendsearch_Click()

framesearch.Visible = False
cmdendsearch.Visible = False
cmdsearch.Visible = True
Form_Load
Timer1.Enabled = False
framefunctiom.Height = 10300
End Sub

Private Sub cmdclose_click()
Unload Me
MDIForm1.Show
End Sub

Private Sub distgrid_Click()
    frmdistdetail.Enabled = False

''' FOR ADDING DATA FROM GRID TO TEXTBOX
    txtdistid.Text = distgrid.TextMatrix(distgrid.Row, 1)
    txtdistname.Text = distgrid.TextMatrix(distgrid.Row, 2)
    txtdistaddress.Text = distgrid.TextMatrix(distgrid.Row, 3)
    txtdistcname.Text = distgrid.TextMatrix(distgrid.Row, 4)
    txtdistcaddress.Text = distgrid.TextMatrix(distgrid.Row, 5)
    txtdistmobileno.Text = distgrid.TextMatrix(distgrid.Row, 6)
End Sub

Private Sub Form_Load()
Call connection
Set rs = New ADODB.Recordset
sql = "select * from DistributorDetail"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic

framesearch.Visible = False
frmdistdetail.Enabled = False
If cmdadd.Enabled = True Then
cmdsave.Enabled = False
End If

framefunctiom.Height = 1
Timer1.Interval = 1
Timer1.Enabled = True

gridload
End Sub

Private Sub gridload()

''' TO GENERATE DISTRIBUTOR GRID
distgrid.TextMatrix(0, 0) = "No."
distgrid.TextMatrix(0, 1) = "Distributor ID"
distgrid.TextMatrix(0, 2) = "Full Name"
distgrid.TextMatrix(0, 3) = "Address"
distgrid.TextMatrix(0, 4) = "Company Name"
distgrid.TextMatrix(0, 5) = "Company Address"
distgrid.TextMatrix(0, 6) = "Mobile"

distgrid.ColWidth(0) = 350
distgrid.ColWidth(1) = 950
distgrid.ColWidth(2) = 1800
distgrid.ColWidth(3) = 3200
distgrid.ColWidth(4) = 1900
distgrid.ColWidth(5) = 3200
distgrid.ColWidth(6) = 1750

''' FOR ADDING RECORDS FROM DATABASE TO GRID
distgrid.Rows = 1
ff = 1
Do While Not rs.EOF
    distgrid.Rows = distgrid.Rows + 1
    distgrid.TextMatrix(ff, 0) = ff
    distgrid.TextMatrix(ff, 1) = IIf(IsNull(rs.Fields("Distributor ID")), "", rs.Fields("Distributor ID"))
    distgrid.TextMatrix(ff, 2) = IIf(IsNull(rs.Fields("Full Name")), "", rs.Fields("Full Name"))
    distgrid.TextMatrix(ff, 3) = IIf(IsNull(rs.Fields("Address")), "", rs.Fields("Address"))
    distgrid.TextMatrix(ff, 4) = IIf(IsNull(rs.Fields("Company Name")), "", rs.Fields("Company Name"))
    distgrid.TextMatrix(ff, 5) = IIf(IsNull(rs.Fields("Company Address")), "", rs.Fields("Company Address"))
    distgrid.TextMatrix(ff, 6) = IIf(IsNull(rs.Fields("Mobile")), "", rs.Fields("Mobile"))
    ff = ff + 1
rs.MoveNext
Loop

End Sub

Private Sub Label1_Click()
framefunctiom.Height = 1
Timer1.Interval = 1
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
framefunctiom.Height = framefunctiom.Height + 1000
If framefunctiom.Height >= 10300 Then
framefunctiom.Height = 10300
Timer1.Enabled = False
End If
End Sub

Private Sub cmdclosemenu_Click()
framefunctiom.Height = 6700
Timer2.Interval = 1
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
framefunctiom.Height = framefunctiom.Height - 100
If framefunctiom.Height <= 350 Then
Timer2.Enabled = False
End If
End Sub

Private Sub txtdistmobileno_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
    If KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
    End If
End Sub

Private Sub txtdistname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
Else
If KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End If
End Sub

Private Sub txtsdistcname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    x = "Company Name"
    y = txtsdistcname.Text
    search
End If
End Sub

Private Sub txtsdistid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    x = "Distributor ID"
    y = txtsdistid.Text
    search
End If
End Sub


Private Sub txtsdistname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
ElseIf KeyAscii = 8 Then
ElseIf KeyAscii = 13 Then
    x = "Full Name"
    y = txtsdistname.Text
search
Else
KeyAscii = 0
End If
End Sub

Private Sub txtsdistmobile_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf KeyAscii = 8 Then
search
ElseIf KeyAscii = 13 Then
    x = "Mobile"
    y = txtsdistmobile.Text
search
Else
KeyAscii = 0
End If
End Sub

Private Sub search()
distgrid.Clear
i = 1
j = 0
rs.MoveFirst
    While rs.EOF = False
    If rs.Fields(x) = y Then

    distgrid.Rows = distgrid.Rows + 1
    distgrid.TextMatrix(0, 0) = "ID"
    distgrid.TextMatrix(0, 1) = "Distributor ID"
    distgrid.TextMatrix(0, 2) = "Full Name"
    distgrid.TextMatrix(0, 3) = "Address"
    distgrid.TextMatrix(0, 4) = "Company Name"
    distgrid.TextMatrix(0, 5) = "Company Address"
    distgrid.TextMatrix(0, 6) = "Mobile"
    j = j + 1
    Call sgrid
    i = i + 1
    Else
    End If
rs.MoveNext
Wend
End Sub

Private Sub sgrid()
    txtdistid.Text = rs.Fields("Distributor ID")
    txtdistname.Text = rs.Fields("Full Name")
    txtdistaddress.Text = rs.Fields("Address")
    txtdistcname.Text = rs.Fields("Company Name")
    txtdistcaddress.Text = rs.Fields("Company Address")
    txtdistmobileno.Text = rs.Fields("Mobile")
    
    distgrid.Rows = i + 1
    distgrid.Row = j
    distgrid.TextMatrix(distgrid.Row, 0) = distgrid.Row
    distgrid.TextMatrix(distgrid.Row, 1) = rs.Fields("distributor ID")
    distgrid.TextMatrix(distgrid.Row, 2) = rs.Fields("Full Name")
    distgrid.TextMatrix(distgrid.Row, 3) = rs.Fields("Address")
    distgrid.TextMatrix(distgrid.Row, 4) = rs.Fields("Company Name")
    distgrid.TextMatrix(distgrid.Row, 5) = rs.Fields("Company Address")
    distgrid.TextMatrix(distgrid.Row, 6) = rs.Fields("Mobile")
End Sub

