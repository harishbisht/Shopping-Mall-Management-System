VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Customer Details"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form5"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   7800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\sms\hr.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\sms\hr.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin vkUserContolsXP.vkFrame framesearch 
      Height          =   2535
      Left            =   4440
      TabIndex        =   31
      Top             =   7800
      Width           =   5655
      _ExtentX        =   9975
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
         Height          =   2055
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3625
         TabCaption(0)   =   "Customer ID"
         TabContCtrlCnt(0)=   4
         Tab(0)ContCtrlCap(1)=   "cmdok"
         Tab(0)ContCtrlCap(2)=   "vkLabel9"
         Tab(0)ContCtrlCap(3)=   "txtscustid"
         Tab(0)ContCtrlCap(4)=   "vkLabel1"
         TabCaption(1)   =   "Customer Name"
         TabContCtrlCnt(1)=   3
         Tab(1)ContCtrlCap(1)=   "XPButton1"
         Tab(1)ContCtrlCap(2)=   "txtscustname"
         Tab(1)ContCtrlCap(3)=   "vkLabel7"
         TabCaption(2)   =   "Mobile Number"
         TabContCtrlCnt(2)=   3
         Tab(2)ContCtrlCap(1)=   "XPButton2"
         Tab(2)ContCtrlCap(2)=   "txtscustmobile"
         Tab(2)ContCtrlCap(3)=   "vkLabel8"
         ActiveTab       =   2
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
         Begin ButtonXp.XPButton XPButton2 
            Height          =   615
            Left            =   2040
            TabIndex        =   42
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
         Begin ButtonXp.XPButton XPButton1 
            Height          =   375
            Left            =   -72480
            TabIndex        =   41
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
         Begin ButtonXp.XPButton cmdok 
            Height          =   255
            Left            =   -71520
            TabIndex        =   39
            Top             =   1560
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
         Begin vkUserContolsXP.vkLabel vkLabel9 
            Height          =   255
            Left            =   -73800
            TabIndex        =   38
            Top             =   480
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Customer ID in the form of ""Cust1"""
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
         Begin vkUserContolsXP.vkTextBox txtscustmobile 
            Height          =   375
            Left            =   2880
            TabIndex        =   37
            Top             =   720
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
            Left            =   600
            TabIndex        =   36
            Top             =   720
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
         Begin vkUserContolsXP.vkTextBox txtscustname 
            Height          =   375
            Left            =   -72120
            TabIndex        =   35
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
         Begin vkUserContolsXP.vkLabel vkLabel7 
            Height          =   375
            Left            =   -74520
            TabIndex        =   34
            Top             =   720
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Enter Customer Name  "
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
         Begin vkUserContolsXP.vkTextBox txtscustid 
            Height          =   375
            Left            =   -72000
            TabIndex        =   33
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Text            =   "Cust"
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
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Left            =   -74040
            TabIndex        =   32
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
      End
   End
   Begin VB.Timer Timer2 
      Left            =   14280
      Top             =   9960
   End
   Begin VB.Timer Timer1 
      Left            =   14280
      Top             =   10440
   End
   Begin vkUserContolsXP.vkFrame framefunctiom 
      Height          =   10590
      Left            =   12960
      TabIndex        =   29
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   18680
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
      Begin ButtonXp.XPButton cmdendsearch 
         Height          =   495
         Left            =   120
         TabIndex        =   18
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
      Begin ButtonXp.XPButton cmdsave 
         Height          =   495
         Left            =   120
         TabIndex        =   9
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
      Begin ButtonXp.XPButton cmdclosemenu 
         Height          =   375
         Left            =   120
         TabIndex        =   17
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
      Begin ButtonXp.XPButton cmdsearch 
         Height          =   495
         Left            =   120
         TabIndex        =   15
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
      Begin ButtonXp.XPButton cmdclose 
         Height          =   495
         Left            =   120
         TabIndex        =   16
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
      Begin ButtonXp.XPButton cmdprevious 
         Height          =   495
         Left            =   120
         TabIndex        =   12
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
      Begin ButtonXp.XPButton cmdnext 
         Height          =   495
         Left            =   120
         TabIndex        =   11
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
      Begin ButtonXp.XPButton cmdbottom 
         Height          =   495
         Left            =   120
         TabIndex        =   13
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
      Begin ButtonXp.XPButton cmdtop 
         Height          =   495
         Left            =   120
         TabIndex        =   10
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
      Begin ButtonXp.XPButton cmdadd 
         Height          =   495
         Left            =   120
         TabIndex        =   5
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
      Begin ButtonXp.XPButton cmdedit 
         Height          =   495
         Left            =   120
         TabIndex        =   6
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
      Begin ButtonXp.XPButton cmddelete 
         Height          =   495
         Left            =   120
         TabIndex        =   7
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
      Begin ButtonXp.XPButton cmdcancel 
         Height          =   495
         Left            =   120
         TabIndex        =   8
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
      Begin VB.Line Line12 
         X1              =   0
         X2              =   1440
         Y1              =   9720
         Y2              =   9720
      End
      Begin VB.Line Line11 
         X1              =   0
         X2              =   1320
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   1320
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   1320
         Y1              =   -240
         Y2              =   -240
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   1320
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   1320
         Y1              =   -360
         Y2              =   -360
      End
      Begin VB.Line Line6 
         X1              =   -240
         X2              =   1080
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1320
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   1320
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   1320
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   1320
         Y1              =   9840
         Y2              =   9840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid custgrid 
      Height          =   5055
      Left            =   120
      TabIndex        =   28
      Top             =   2760
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin vkUserContolsXP.vkFrame frmcustdetail 
      Height          =   2415
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   4260
      Caption         =   "Customer Details"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComCtl2.DTPicker txtbirthdate 
         Height          =   375
         Left            =   8760
         TabIndex        =   40
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   39548
      End
      Begin vkUserContolsXP.vkLabel vkLabel6 
         Height          =   375
         Left            =   8760
         TabIndex        =   27
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Birthdate"
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
      Begin vkUserContolsXP.vkLabel vkLabel5 
         Height          =   375
         Left            =   10440
         TabIndex        =   26
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
      Begin vkUserContolsXP.vkTextBox txtcustmobileno 
         Height          =   375
         Left            =   10440
         TabIndex        =   4
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   375
         Left            =   7800
         TabIndex        =   25
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Gender"
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
      Begin vkUserContolsXP.vkTextBox txtcustname 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   960
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   375
         Left            =   1560
         TabIndex        =   24
         Top             =   600
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Full Name"
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
         Left            =   3720
         TabIndex        =   23
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
      Begin vkUserContolsXP.vkTextBox txtcustid 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkOptionButton optmale 
         Height          =   255
         Left            =   7800
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Male"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton optfemale 
         Height          =   255
         Left            =   7800
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Female"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkTextBox txtcustaddress 
         Height          =   855
         Left            =   3720
         TabIndex        =   1
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkLabel Cust 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Customer ID"
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
End
Attribute VB_Name = "Form5"
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
frmcustdetail.Enabled = True
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

'''TO GENERATE AUTOMATIC CUSTOMER ID
    If custgrid.Rows = 1 Then
        txtcustid.Text = "Cust1"
'    ElseIf custgrid.Rows = 2 Then
'        txtcustid.Text = "Cust2"
    Else
        txtcustid.Text = "Cust" & (custgrid.Rows - 1 + 1)
        txtcustid.Enabled = False
    End If
    txtcustname.SetFocus
'''--------------------------------------------------------------------------
End Sub

Private Sub cmdbottom_Click()
custgrid.Row = custgrid.Rows - 1
custgrid_Click
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

Private Sub cmdclear_Click()
txtcustid.Text = ""
txtcustname.Text = ""
txtcustaddress.Text = ""
optmale.Value = False And optfemale.Value = False
txtbirthdate.Value = Now
txtcustmobileno.Text = ""
End Sub

Private Sub cmdclose_click()
MDIForm1.Show
Unload Me
End Sub

Private Sub cmdclosemenu_Click()
framefunctiom.Height = 6700
Timer2.Interval = 1
Timer2.Enabled = True
End Sub

Private Sub cmddelete_Click()
If custgrid.Rows = 1 Then
MsgBox "No Record found in Database"
Exit Sub
Else
click = MsgBox("Do u want to Delete this Record from database?", vbYesNo + vbQuestion)
If click = vbYes Then
rs.MoveFirst
While rs.EOF = False
If rs.Fields("Customer ID") = txtcustid.Text Then
rs.Delete
custgrid.Clear
'custgrid.Rows = 1
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
cmdclear_Click
End Sub

Private Sub cmdedit_Click()
frmcustdetail.Enabled = True
addflag = False
cmdsave.Enabled = True
txtcustid.Enabled = True
txtcustname.Enabled = True
txtcustaddress.Enabled = True
txtcustmobileno.Enabled = True
txtbirthdate.Enabled = True
End Sub

Private Sub cmdendsearch_Click()
Form_Load
Timer1.Enabled = False
framefunctiom.Height = 10300
framesearch.Visible = False
cmdendsearch.Visible = False
cmdsearch.Visible = True
End Sub

Private Sub cmdnext_Click()
If custgrid.Row < custgrid.Rows - 1 Then
   custgrid.Row = custgrid.Row + 1
custgrid_Click
   cmdnext.SetFocus
End If
End Sub

Private Sub sgrid()
txtcustid.Text = rs.Fields("Customer ID")
txtcustname.Text = rs.Fields("Full Name")
txtcustaddress.Text = rs.Fields("Address")
    If rs.Fields("Gender") = "Male" Then
    optmale.Value = vbChecked
    Else
    optfemale.Value = vbChecked
    End If
txtbirthdate.Value = rs.Fields("Birthdate")
txtcustmobileno.Text = rs.Fields("Mobile")
    custgrid.Rows = i + 1
    custgrid.Row = j
    custgrid.TextMatrix(custgrid.Row, 0) = custgrid.Row
    custgrid.TextMatrix(custgrid.Row, 1) = rs.Fields("Customer ID")
    custgrid.TextMatrix(custgrid.Row, 2) = rs.Fields("Full Name")
    custgrid.TextMatrix(custgrid.Row, 3) = rs.Fields("Address")
    custgrid.TextMatrix(custgrid.Row, 4) = rs.Fields("Gender")
    custgrid.TextMatrix(custgrid.Row, 5) = rs.Fields("Birthdate")
    custgrid.TextMatrix(custgrid.Row, 6) = rs.Fields("Mobile")
End Sub


Private Sub cmdprevious_Click()
If custgrid.Row > 1 Then
    If custgrid.Row - 1 < custgrid.Rows Then
        custgrid.Row = custgrid.Row - 1
    custgrid_Click
        cmdprevious.SetFocus
    End If
End If
End Sub

Private Sub cmdsave_Click()
If txtcustname.Text = "" Or txtcustaddress.Text = "" Or (optmale.Value = False And optfemale.Value = False) Or txtcustmobileno.Text = "" Then
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
        rs.Fields("Customer ID") = txtcustid.Text
        rs.Fields("Full Name") = txtcustname.Text
        rs.Fields("Address") = txtcustaddress.Text
    If optmale.Value = vbChecked Then
        rs.Fields("Gender") = optmale.Caption
    ElseIf optfemale.Value = vbChecked Then
        rs.Fields("Gender") = "Female"
        End If
        rs.Fields("Birthdate") = txtbirthdate.Value
        rs.Fields("Mobile") = txtcustmobileno.Text
        rs.Update
    Call gridfill

Else
''''' FOR EDITING RECORDS IN DATABASE AND GRID
        rs.MoveFirst
        Do While Not rs.EOF
         If rs.Fields("Customer Id") = txtcustid.Text Then
        Exit Do
         End If
        rs.MoveNext
        Loop
        
''' TO EDIT DATABASE
        rs.Fields("Customer ID") = txtcustid.Text
        rs.Fields("Full Name") = txtcustname.Text
        rs.Fields("Address") = txtcustaddress.Text
    If optmale.Value = vbChecked Then
        rs.Fields("Gender") = optmale.Caption
    ElseIf optfemale.Value = vbChecked Then
        rs.Fields("Gender") = "Female"
    End If
        rs.Fields("Birthdate") = txtbirthdate.Value
        rs.Fields("Mobile") = txtcustmobileno.Text
        rs.Update
        
''' TO FILL GRID FROM DATABASE
    custgrid.TextMatrix(custgrid.Row, 2) = rs.Fields("Full Name")
    custgrid.TextMatrix(custgrid.Row, 3) = rs.Fields("Address")
    custgrid.TextMatrix(custgrid.Row, 4) = rs.Fields("Gender")
    custgrid.TextMatrix(custgrid.Row, 5) = rs.Fields("Birthdate")
    custgrid.TextMatrix(custgrid.Row, 6) = rs.Fields("Mobile")
End If
cmdclear_Click
End Sub

Private Sub cmdsearch_Click()
framesearch.Visible = True
cmdsearch.Visible = False
cmdendsearch.Visible = True
End Sub

Private Sub cmdtop_Click()
    custgrid.Row = 1
    custgrid.SetFocus
custgrid_Click
    End Sub

Private Sub custgrid_Click()
frmcustdetail.Enabled = False
frmcustdetail.Enabled = False
''' FOR ADDING DATA FROM GRID TO TEXTBOX
        txtcustid.Text = custgrid.TextMatrix(custgrid.Row, 1)
        txtcustname.Text = custgrid.TextMatrix(custgrid.Row, 2)
        txtcustaddress.Text = custgrid.TextMatrix(custgrid.Row, 3)
'''for Gender selection
        If custgrid.TextMatrix(custgrid.Row, 4) = "Male" Then
          optmale.Value = vbChecked
        ElseIf custgrid.TextMatrix(custgrid.Row, 4) = "Female" Then
          optfemale.Value = vbChecked
        End If
'''for empty grid click
        If custgrid.Rows = 1 Then
        txtbirthdate.Value = Now
        Else
        txtbirthdate.Value = custgrid.TextMatrix(custgrid.Row, 5)
        txtcustmobileno.Text = custgrid.TextMatrix(custgrid.Row, 6)
        End If
End Sub

Private Sub Form_Load()
Call connection
Set rs = New ADODB.Recordset
sql = "select * from CustomerDetail"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic

framesearch.Visible = False
frmcustdetail.Enabled = False
If cmdadd.Enabled = True Then
cmdsave.Enabled = False
End If

framefunctiom.Height = 1
Timer1.Interval = 1
Timer1.Enabled = True

gridload
End Sub
Private Sub gridload()

''' TO GENERATE CUSTOMER GRID
custgrid.TextMatrix(0, 0) = "ID"
custgrid.TextMatrix(0, 1) = "Customer ID"
custgrid.TextMatrix(0, 2) = "Full Name"
custgrid.TextMatrix(0, 3) = "Address"
custgrid.TextMatrix(0, 4) = "Gender"
custgrid.TextMatrix(0, 5) = "Birthdate"
custgrid.TextMatrix(0, 6) = "Mobile"

custgrid.ColWidth(0) = 400
custgrid.ColWidth(1) = 1000
custgrid.ColWidth(2) = 2200
custgrid.ColWidth(3) = 4100
custgrid.ColWidth(4) = 1200
custgrid.ColWidth(5) = 1600
custgrid.ColWidth(6) = 1900

''' FOR ADDING RECORDS FROM DATABASE TO GRID
custgrid.Rows = 1
ff = 1
Do While Not rs.EOF
    custgrid.Rows = custgrid.Rows + 1
    custgrid.TextMatrix(ff, 0) = ff
    custgrid.TextMatrix(ff, 1) = IIf(IsNull(rs.Fields("customer ID")), "", rs.Fields("customer ID"))
    custgrid.TextMatrix(ff, 2) = IIf(IsNull(rs.Fields("Full Name")), "", rs.Fields("Full Name"))
    custgrid.TextMatrix(ff, 3) = IIf(IsNull(rs.Fields("Address")), "", rs.Fields("Address"))
    custgrid.TextMatrix(ff, 4) = IIf(IsNull(rs.Fields("Gender")), "", rs.Fields("Gender"))
    custgrid.TextMatrix(ff, 5) = IIf(IsNull(rs.Fields("Birthdate")), "", rs.Fields("Birthdate"))
    custgrid.TextMatrix(ff, 6) = IIf(IsNull(rs.Fields("Mobile")), "", rs.Fields("Mobile"))
    ff = ff + 1
rs.MoveNext
Loop

End Sub

Private Sub gridfill()
custgrid.Rows = custgrid.Rows + 1
If custgrid.Rows = 2 Then
custgrid.TextMatrix(custgrid.Rows - 1, 0) = 1
Else
custgrid.TextMatrix(custgrid.Rows - 1, 0) = custgrid.TextMatrix(custgrid.Rows - 2, 0) + 1
End If
    
    custgrid.TextMatrix(custgrid.Rows - 1, 1) = txtcustid.Text
    custgrid.TextMatrix(custgrid.Rows - 1, 2) = txtcustname.Text
    custgrid.TextMatrix(custgrid.Rows - 1, 3) = txtcustaddress.Text
 
         '''FOR GENDER SELECTION
        If optmale.Value = vbChecked Then
            custgrid.TextMatrix(custgrid.Rows - 1, 4) = "Male"
        ElseIf optfemale.Value = vbChecked Then
            custgrid.TextMatrix(custgrid.Rows - 1, 4) = "Female"
        Else
        End If
    custgrid.TextMatrix(custgrid.Rows - 1, 5) = txtbirthdate.Value
    custgrid.TextMatrix(custgrid.Rows - 1, 6) = txtcustmobileno.Text
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

Private Sub Timer2_Timer()
framefunctiom.Height = framefunctiom.Height - 100
If framefunctiom.Height <= 350 Then
Timer2.Enabled = False
End If
End Sub

Private Sub txtcustmobileno_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
    If KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
    End If
End Sub

Private Sub txtcustname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
Else
If KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End If
End Sub

Private Sub txtscustid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    x = "Customer ID"
    y = txtscustid.Text
    search
End If
End Sub


Private Sub txtscustname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
ElseIf KeyAscii = 8 Then
ElseIf KeyAscii = 13 Then
    x = "Full Name"
    y = txtscustname.Text
search
Else
KeyAscii = 0
End If
End Sub

Private Sub txtscustmobile_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf KeyAscii = 8 Then
search
ElseIf KeyAscii = 13 Then
    x = "Mobile"
    y = txtscustmobile.Text
search
Else
KeyAscii = 0
End If
End Sub

'Private Sub XPButton1_Click()
'x = "Full Name"
'y = txtscustname.Text
'search
'End Sub

'Private Sub XPButton2_Click()
'x = "Mobile"
'y = txtscustmobile.Text
'search
'End Sub

Private Sub XTab1_BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
txtscustid.Text = "Cust"
txtscustname.Text = ""
txtscustmobile.Text = ""
End Sub

Private Sub search()
custgrid.Clear
i = 1
j = 0
rs.MoveFirst
    While rs.EOF = False
    If rs.Fields(x) = y Then

    custgrid.Rows = custgrid.Rows + 1
    custgrid.TextMatrix(0, 0) = "ID"
    custgrid.TextMatrix(0, 1) = "Customer ID"
    custgrid.TextMatrix(0, 2) = "Full Name"
    custgrid.TextMatrix(0, 3) = "Address"
    custgrid.TextMatrix(0, 4) = "Gender"
    custgrid.TextMatrix(0, 5) = "Birthdate"
    custgrid.TextMatrix(0, 6) = "Mobile"
    j = j + 1

    Call sgrid
    i = i + 1
    Else
    End If
rs.MoveNext
Wend
End Sub
