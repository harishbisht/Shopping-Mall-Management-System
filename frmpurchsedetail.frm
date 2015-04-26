VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form8"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin vkUserContolsXP.vkTextBox txtseldistadd 
      Height          =   375
      Left            =   9240
      TabIndex        =   35
      Top             =   480
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
   Begin vkUserContolsXP.vkTextBox txtseldistname 
      Height          =   375
      Left            =   7680
      TabIndex        =   34
      Top             =   480
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
   Begin vkUserContolsXP.vkTextBox txtseldistid 
      Height          =   375
      Left            =   6240
      TabIndex        =   32
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   6480
      TabIndex        =   31
      Top             =   8280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   2
      Format          =   "##"
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin vkUserContolsXP.vkLabel vkLabel5 
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   8280
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "%"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   3135
      Left            =   5880
      TabIndex        =   10
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5530
      Caption         =   "Disributers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ListView ListView1 
         Height          =   1875
         Left            =   240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3307
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Dist Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Company Name"
            Object.Width           =   3069
         EndProperty
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   2535
      Left            =   5400
      TabIndex        =   9
      Top             =   5040
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4471
      Caption         =   "Total Purchase"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSFlexGridLib.MSFlexGrid purchgrid 
         Height          =   1965
         Left            =   240
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   480
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   3466
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ForeColor       =   0
         Enabled         =   -1  'True
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   6120
      TabIndex        =   8
      Top             =   4200
      Width           =   5415
   End
   Begin vkUserContolsXP.vkLabel vkLabel4 
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   9120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Grant Total"
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
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   8280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Discounts"
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
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   7800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Amounts"
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
   Begin vkUserContolsXP.vkTextBox vkTextBox4 
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   9120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
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
   Begin vkUserContolsXP.vkTextBox vkTextBox2 
      Height          =   300
      Left            =   6480
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
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
   Begin vkUserContolsXP.vkTextBox vkTextBox1 
      Height          =   300
      Left            =   2760
      TabIndex        =   2
      Top             =   8040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
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
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   8040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      BackColor       =   16777215
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
   Begin prjXTab.XTab XTab1 
      Height          =   6735
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11880
      TabCount        =   4
      TabCaption(0)   =   "General"
      TabContCtrlCnt(0)=   3
      Tab(0)ContCtrlCap(1)=   "vkLabel35"
      Tab(0)ContCtrlCap(2)=   "frametoothpaste"
      Tab(0)ContCtrlCap(3)=   "cmbgeneralitem"
      TabCaption(1)   =   "Clothes"
      TabContCtrlCnt(1)=   2
      Tab(1)ContCtrlCap(1)=   "vkFrame3"
      Tab(1)ContCtrlCap(2)=   "cmbclotheitem"
      TabCaption(2)   =   "Electronic"
      TabContCtrlCnt(2)=   2
      Tab(2)ContCtrlCap(1)=   "vkFrame4"
      Tab(2)ContCtrlCap(2)=   "cmbelectitem"
      TabCaption(3)   =   "Footwear"
      TabContCtrlCnt(3)=   3
      Tab(3)ContCtrlCap(1)=   "lbl4"
      Tab(3)ContCtrlCap(2)=   "vkFrame5"
      Tab(3)ContCtrlCap(3)=   "cmbfootitem"
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
      Begin vkUserContolsXP.vkLabel lbl4 
         Height          =   255
         Left            =   -72960
         TabIndex        =   87
         Top             =   5280
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "..Invalid.."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin vkUserContolsXP.vkLabel vkLabel35 
         Height          =   375
         Left            =   2400
         TabIndex        =   82
         Top             =   4320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkFrame vkFrame5 
         Height          =   4935
         Left            =   -74520
         TabIndex        =   66
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   8705
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   16711680
         TitleColor2     =   16777215
         BorderColor     =   16711680
         DisplayPicture  =   0   'False
         Begin vkUserContolsXP.vkLabel vkLabel33 
            Height          =   375
            Left            =   1920
            TabIndex        =   80
            Top             =   2760
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "%"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vkUserContolsXP.vkTextBox txtproddiscf 
            Height          =   375
            Left            =   1320
            TabIndex        =   79
            Top             =   2760
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
         Begin vkUserContolsXP.vkTextBox txtprodstockf 
            Height          =   375
            Left            =   1320
            TabIndex        =   78
            Top             =   2160
            Width           =   975
            _ExtentX        =   1720
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
         Begin vkUserContolsXP.vkTextBox txtprodpricef 
            Height          =   375
            Left            =   1320
            TabIndex        =   77
            Top             =   1560
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
            LegendForeColor =   12937777
         End
         Begin vkUserContolsXP.vkTextBox txtprodnamef 
            Height          =   375
            Left            =   1320
            TabIndex        =   76
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
            BorderColor     =   16711680
            LegendForeColor =   255
         End
         Begin vkUserContolsXP.vkLabel vkLabel32 
            Height          =   375
            Left            =   360
            TabIndex        =   75
            Top             =   2880
            Width           =   855
            _ExtentX        =   1508
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
         Begin vkUserContolsXP.vkLabel vkLabel31 
            Height          =   375
            Left            =   360
            TabIndex        =   74
            Top             =   2280
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel30 
            Height          =   375
            Left            =   360
            TabIndex        =   73
            Top             =   1680
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel29 
            Height          =   375
            Left            =   360
            TabIndex        =   72
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkTextBox txtprodqtyf 
            Height          =   375
            Left            =   1320
            TabIndex        =   71
            Top             =   3720
            Width           =   975
            _ExtentX        =   1720
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
         Begin vkUserContolsXP.vkLabel vkLabel28 
            Height          =   375
            Left            =   360
            TabIndex        =   70
            Top             =   3720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Quantity"
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
         Begin ButtonXp.XPButton cmdfpurchase 
            Height          =   495
            Left            =   1200
            TabIndex        =   69
            Top             =   4320
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
            Caption         =   "Purchase"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin vkUserContolsXP.vkTextBox txtprodidf 
            Height          =   375
            Left            =   1320
            TabIndex        =   68
            Top             =   360
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
            BorderColor     =   16711680
            LegendForeColor =   255
         End
         Begin vkUserContolsXP.vkLabel vkLabel27 
            Height          =   375
            Left            =   360
            TabIndex        =   67
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
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
         Begin VB.Line Line5 
            X1              =   120
            X2              =   3360
            Y1              =   3480
            Y2              =   3480
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame4 
         Height          =   4935
         Left            =   -74520
         TabIndex        =   51
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   8705
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   16711680
         TitleColor2     =   16777215
         BorderColor     =   16711680
         DisplayPicture  =   0   'False
         Begin vkUserContolsXP.vkLabel lbl3 
            Height          =   255
            Left            =   1560
            TabIndex        =   86
            Top             =   3720
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "..Invalid.."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
         Begin vkUserContolsXP.vkLabel vkLabel34 
            Height          =   375
            Left            =   1920
            TabIndex        =   81
            Top             =   2760
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "%"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vkUserContolsXP.vkLabel vkLabel26 
            Height          =   375
            Left            =   360
            TabIndex        =   65
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
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
         Begin vkUserContolsXP.vkTextBox txtprodide 
            Height          =   375
            Left            =   1320
            TabIndex        =   64
            Top             =   360
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
            BorderColor     =   16711680
            LegendForeColor =   255
         End
         Begin ButtonXp.XPButton cmdepurchase 
            Height          =   495
            Left            =   1200
            TabIndex        =   63
            Top             =   4320
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
            Caption         =   "Purchase"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin vkUserContolsXP.vkLabel vkLabel25 
            Height          =   375
            Left            =   360
            TabIndex        =   62
            Top             =   3720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Quantity"
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
         Begin vkUserContolsXP.vkTextBox txtprodqtye 
            Height          =   375
            Left            =   1320
            TabIndex        =   61
            Top             =   3720
            Width           =   975
            _ExtentX        =   1720
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
         Begin vkUserContolsXP.vkLabel vkLabel24 
            Height          =   375
            Left            =   360
            TabIndex        =   60
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel23 
            Height          =   375
            Left            =   360
            TabIndex        =   59
            Top             =   1680
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel22 
            Height          =   375
            Left            =   360
            TabIndex        =   58
            Top             =   2280
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel21 
            Height          =   375
            Left            =   360
            TabIndex        =   57
            Top             =   2880
            Width           =   855
            _ExtentX        =   1508
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
         Begin vkUserContolsXP.vkTextBox txtprodnamee 
            Height          =   375
            Left            =   1320
            TabIndex        =   56
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
            BorderColor     =   16711680
            LegendForeColor =   255
         End
         Begin vkUserContolsXP.vkTextBox txtprodpricee 
            Height          =   375
            Left            =   1320
            TabIndex        =   55
            Top             =   1560
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
            LegendForeColor =   12937777
         End
         Begin vkUserContolsXP.vkTextBox txtprodstocke 
            Height          =   375
            Left            =   1320
            TabIndex        =   54
            Top             =   2160
            Width           =   975
            _ExtentX        =   1720
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
         Begin vkUserContolsXP.vkTextBox txtproddisce 
            Height          =   375
            Left            =   1320
            TabIndex        =   53
            Top             =   2760
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
         Begin vkUserContolsXP.vkLabel vkLabel20 
            Height          =   375
            Left            =   2040
            TabIndex        =   52
            Top             =   2280
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "%"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   3360
            Y1              =   3480
            Y2              =   3480
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame3 
         Height          =   4935
         Left            =   -74520
         TabIndex        =   36
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   8705
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   16711680
         TitleColor2     =   16777215
         BorderColor     =   16711680
         DisplayPicture  =   0   'False
         Begin vkUserContolsXP.vkLabel lbl2 
            Height          =   255
            Left            =   1560
            TabIndex        =   85
            Top             =   3720
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "..Invalid.."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
         Begin vkUserContolsXP.vkLabel vkLabel19 
            Height          =   375
            Left            =   1920
            TabIndex        =   50
            Top             =   2760
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "%"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vkUserContolsXP.vkTextBox txtproddiscc 
            Height          =   375
            Left            =   1320
            TabIndex        =   49
            Top             =   2760
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
         Begin vkUserContolsXP.vkTextBox txtprodstockc 
            Height          =   375
            Left            =   1320
            TabIndex        =   48
            Top             =   2160
            Width           =   975
            _ExtentX        =   1720
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
         Begin vkUserContolsXP.vkTextBox txtprodpricec 
            Height          =   375
            Left            =   1320
            TabIndex        =   47
            Top             =   1560
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
            LegendForeColor =   12937777
         End
         Begin vkUserContolsXP.vkTextBox txtprodnamec 
            Height          =   375
            Left            =   1320
            TabIndex        =   46
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
            BorderColor     =   16711680
            LegendForeColor =   255
         End
         Begin vkUserContolsXP.vkLabel vkLabel18 
            Height          =   375
            Left            =   360
            TabIndex        =   45
            Top             =   2880
            Width           =   855
            _ExtentX        =   1508
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
         Begin vkUserContolsXP.vkLabel vkLabel17 
            Height          =   375
            Left            =   360
            TabIndex        =   44
            Top             =   2280
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel16 
            Height          =   375
            Left            =   360
            TabIndex        =   43
            Top             =   1680
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel15 
            Height          =   375
            Left            =   360
            TabIndex        =   42
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkTextBox txtprodqtyc 
            Height          =   375
            Left            =   1320
            TabIndex        =   41
            Top             =   3720
            Width           =   975
            _ExtentX        =   1720
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
         Begin vkUserContolsXP.vkLabel vkLabel14 
            Height          =   375
            Left            =   360
            TabIndex        =   40
            Top             =   3720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Quantity"
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
         Begin ButtonXp.XPButton cmdcpurchase 
            Height          =   495
            Left            =   1200
            TabIndex        =   39
            Top             =   4320
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
            Caption         =   "Purchase"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin vkUserContolsXP.vkTextBox txtprodidc 
            Height          =   375
            Left            =   1320
            TabIndex        =   38
            Top             =   360
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
            BorderColor     =   16711680
            LegendForeColor =   255
         End
         Begin vkUserContolsXP.vkLabel vkLabel13 
            Height          =   375
            Left            =   360
            TabIndex        =   37
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
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
         Begin VB.Line Line3 
            X1              =   120
            X2              =   3360
            Y1              =   3480
            Y2              =   3480
         End
      End
      Begin VB.ComboBox cmbfootitem 
         Height          =   315
         Left            =   -74520
         TabIndex        =   28
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cmbelectitem 
         Height          =   315
         Left            =   -74520
         TabIndex        =   27
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cmbclotheitem 
         Height          =   315
         Left            =   -74520
         TabIndex        =   26
         Top             =   840
         Width           =   2655
      End
      Begin vkUserContolsXP.vkFrame frametoothpaste 
         Height          =   4935
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   8705
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   16711680
         TitleColor2     =   16777215
         BorderColor     =   16711680
         DisplayPicture  =   0   'False
         Begin vkUserContolsXP.vkLabel lbl1 
            Height          =   255
            Left            =   1560
            TabIndex        =   83
            Top             =   3720
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "..Invalid.."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
         Begin vkUserContolsXP.vkLabel vkLabel12 
            Height          =   375
            Left            =   360
            TabIndex        =   30
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
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
         Begin vkUserContolsXP.vkTextBox txtprodidg 
            Height          =   375
            Left            =   1320
            TabIndex        =   29
            Top             =   360
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
            BorderColor     =   16711680
            LegendForeColor =   255
         End
         Begin ButtonXp.XPButton cmdgpurchase 
            Height          =   495
            Left            =   1200
            TabIndex        =   25
            Top             =   4320
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
            Caption         =   "Purchase"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin vkUserContolsXP.vkLabel vkLabel6 
            Height          =   375
            Left            =   360
            TabIndex        =   24
            Top             =   3720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Quantity"
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
         Begin vkUserContolsXP.vkTextBox txtprodqtyg 
            Height          =   375
            Left            =   1320
            TabIndex        =   23
            Top             =   3720
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   360
            TabIndex        =   22
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel8 
            Height          =   375
            Left            =   360
            TabIndex        =   21
            Top             =   1680
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel9 
            Height          =   375
            Left            =   360
            TabIndex        =   20
            Top             =   2280
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel vkLabel10 
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   2880
            Width           =   855
            _ExtentX        =   1508
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
         Begin vkUserContolsXP.vkTextBox txtprodnameg 
            Height          =   375
            Left            =   1320
            TabIndex        =   18
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
            BorderColor     =   16711680
            LegendForeColor =   255
         End
         Begin vkUserContolsXP.vkTextBox txtprodpriceg 
            Height          =   375
            Left            =   1320
            TabIndex        =   17
            Top             =   1560
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
            LegendForeColor =   12937777
         End
         Begin vkUserContolsXP.vkTextBox txtprodstockg 
            Height          =   375
            Left            =   1320
            TabIndex        =   16
            Top             =   2160
            Width           =   975
            _ExtentX        =   1720
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
         Begin vkUserContolsXP.vkTextBox txtproddiscg 
            Height          =   375
            Left            =   1320
            TabIndex        =   15
            Top             =   2760
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
         Begin vkUserContolsXP.vkLabel vkLabel11 
            Height          =   375
            Left            =   2040
            TabIndex        =   14
            Top             =   2280
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "%"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   3360
            Y1              =   3480
            Y2              =   3480
         End
      End
      Begin VB.ComboBox cmbgeneralitem 
         Height          =   315
         ItemData        =   "frmpurchsedetail.frx":0000
         Left            =   480
         List            =   "frmpurchsedetail.frx":0002
         TabIndex        =   12
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.Line Line1 
      X1              =   5040
      X2              =   9120
      Y1              =   8760
      Y2              =   8760
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rsdist As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim ff As Integer
Dim sql As String
Dim str As String
Dim str1 As String
Dim a As String
Dim item, t1, t2, t3, t4, t5
Dim i As Integer


Private Sub cmbgeneralitem_Click()
j = cmbgeneralitem.ListCount
For i = 0 To j
Select Case cmbgeneralitem.ListIndex
Case i
item = cmbgeneralitem.Text
Call data
txtprodidg.Text = t1
txtprodnameg.Text = t2
txtprodpriceg.Text = t3
txtprodstockg.Text = t4
txtproddiscg.Text = t5
End Select
Next
End Sub

Private Sub cmbclotheitem_Click()
j = cmbclotheitem.ListCount
For i = 0 To j
Select Case cmbclotheitem.ListIndex
Case i
item = cmbclotheitem.Text
Call data
txtprodidc.Text = t1
txtprodnamec.Text = t2
txtprodpricec.Text = t3
txtprodstockc.Text = t4
txtproddiscc.Text = t5
End Select
Next
End Sub

Private Sub cmbelectitem_Click()
j = cmbelectitem.ListCount
For i = 0 To j
Select Case cmbelectitem.ListIndex
Case i
item = cmbelectitem.Text
Call data
txtprodide.Text = t1
txtprodnamee.Text = t2
txtprodpricee.Text = t3
txtprodstocke.Text = t4
txtproddisce.Text = t5
End Select
Next
End Sub
Private Sub cmbfootitem_Click()
j = cmbfootitem.ListCount
For i = 0 To j
Select Case cmbfootitem.ListIndex
Case i
item = cmbfootitem.Text
Call data
txtprodidf.Text = t1
txtprodnamef.Text = t2
txtprodpricef.Text = t3
txtprodstockf.Text = t4
txtproddiscf.Text = t5
End Select
Next
End Sub

Private Sub data()
str1 = "select * from ProductDetail"
rs1.Open str1, cn, adOpenDynamic, adLockOptimistic
rs1.MoveFirst
Do While Not rs1.EOF
If rs1.Fields("Prod Name") = item Then
t1 = rs1.Fields("Prod ID")
t2 = rs1.Fields("Prod Name")
t3 = rs1.Fields("Price")
t4 = rs1.Fields("Stock")
t5 = rs1.Fields("Discount")
rs1.MoveNext
Else
rs1.MoveNext
End If
Loop
rs1.Close
End Sub

Private Sub cmdgpurchase_Click()
sg = txtprodstockg.Text
qg = txtprodqtyg.Text
If qg = "" Or qg = 0 Or qg > sg Then
txtprodqtyg.Text = ""
lbl1.Visible = True
End If

Call gridadd

End Sub

Private Sub cmdcpurchase_Click()
sc = txtprodstockc.Text
qc = txtprodqtyc.Text
If qc = "" Or qc = 0 Or qc > sc Then
txtprodqtyc.Text = ""
lbl2.Visible = True
End If
End Sub

Private Sub cmdepurchase_Click()
se = txtprodstocke.Text
qe = txtprodqtye.Text
If qe = "" Or qe = 0 Or qe > se Then
txtprodqtye.Text = ""
lbl3.Visible = True
End If
End Sub

Private Sub cmdfpurchase_Click()
sf = txtprodstockf.Text
qf = txtprodqtyf.Text
If qf = "" Or qf = 0 Or qf > sf Then
txtprodqtyf.Text = ""
lbl4.Visible = True
End If
End Sub

Private Sub gridadd()
''' TO GENERATE PURCHASE GRID
purchgrid.TextMatrix(0, 0) = "No."
purchgrid.TextMatrix(0, 1) = "Prod ID"
purchgrid.TextMatrix(0, 2) = "Prod Name"
purchgrid.TextMatrix(0, 3) = "Company Name"
purchgrid.TextMatrix(0, 4) = "Price"
purchgrid.TextMatrix(0, 5) = "Discount"
purchgrid.TextMatrix(0, 6) = "Quantity"
purchgrid.TextMatrix(0, 7) = "Amount"

purchgrid.ColWidth(0) = 350
purchgrid.ColWidth(1) = 950
purchgrid.ColWidth(2) = 2100
purchgrid.ColWidth(3) = 2100
purchgrid.ColWidth(4) = 700
purchgrid.ColWidth(5) = 800
purchgrid.ColWidth(6) = 1000
purchgrid.ColWidth(7) = 1000

End Sub

Private Sub Form_Load()
Call connection
sql = "select* from DistributorDetail"
rsdist.Open sql, cn, adOpenDynamic, adLockOptimistic
ff = 1
Do While Not rsdist.EOF
    ListView1.ListItems.Add , , rsdist.Fields("Distributor ID")
    ListView1.ListItems(ff).SubItems(1) = rsdist.Fields("Full Name")
    ListView1.ListItems(ff).SubItems(2) = rsdist.Fields("Company Name")
ff = ff + 1
rsdist.MoveNext
Loop
Call gridadd
End Sub

Private Sub ListView1_Click()
    txtseldistid.Text = ListView1.SelectedItem.Text
    txtseldistname.Text = ListView1.SelectedItem.ListSubItems(1).Text
    txtseldistadd.Text = ListView1.SelectedItem.ListSubItems(2).Text
    Call purc
End Sub
'==============================================
Private Sub purc()
cmbgeneralitem.Clear
cmbclotheitem.Clear
cmbelectitem.Clear
cmbfootitem.Clear
str = "select*from ProductDetail"
rs.Open str, cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs.Fields("Dist ID") = txtseldistid.Text Then
If rs.Fields("Type") = "General" Then
a = rs.Fields("Prod Name")
cmbgeneralitem.AddItem (a)
ElseIf rs.Fields("Type") = "Clothes" Then
b = rs.Fields("Prod Name")
cmbclotheitem.AddItem (b)
ElseIf rs.Fields("Type") = "Electronics" Then
c = rs.Fields("Prod Name")
cmbelectitem.AddItem (c)
ElseIf rs.Fields("Type") = "Footware" Then
d = rs.Fields("Prod Name")
cmbfootitem.AddItem (d)
End If
End If
rs.MoveNext
Loop
rs.Close
End Sub

Private Sub txtprodqtyg_Click()
lbl1.Visible = False
End Sub

Private Sub txtprodqtyc_Click()
lbl2.Visible = False
End Sub
Private Sub txtprodqtye_Click()
lbl3.Visible = False
End Sub
Private Sub txtprodqtyf_Click()
lbl4.Visible = False
End Sub


'==============================================
Private Sub txtprodqtyg_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 8 Or KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtprodqtyc_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 8 Or KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtprodqtye_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 8 Or KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtprodqtyf_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 8 Or KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub XTab1_Click()
cmbgeneralitem.Text = ""
cmbclotheitem.Text = ""
cmbelectitem.Text = ""
cmbfootitem.Text = ""

txtprodidg.Text = ""
txtprodidc.Text = ""
txtprodide.Text = ""
txtprodidf.Text = ""

txtprodnameg.Text = ""
txtprodnamec.Text = ""
txtprodnamee.Text = ""
txtprodnamef.Text = ""

txtprodpriceg.Text = ""
txtprodpricec.Text = ""
txtprodpricee.Text = ""
txtprodpricef.Text = ""

txtprodstockg.Text = ""
txtprodstockc.Text = ""
txtprodstocke.Text = ""
txtprodstockf.Text = ""

txtproddiscg.Text = ""
txtproddiscc.Text = ""
txtproddisce.Text = ""
txtproddiscf.Text = ""

End Sub
