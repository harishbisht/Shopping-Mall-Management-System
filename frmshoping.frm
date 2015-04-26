VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   10845
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form7"
   ScaleHeight     =   10845
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin vkUserContolsXP.vkCommand vkCommand1 
      Height          =   615
      Left            =   1320
      TabIndex        =   20
      Top             =   7320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      Caption         =   "New  Customer then click here"
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
      CustomStyle     =   0
   End
   Begin vkUserContolsXP.vkLabel vkLabel3 
      Height          =   255
      Left            =   7080
      TabIndex        =   18
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      BackColor       =   16777215
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
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   1095
      Left            =   4440
      TabIndex        =   5
      Top             =   4440
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowTitle       =   0   'False
      Begin vkUserContolsXP.vkLabel vkLabel8 
         Height          =   255
         Left            =   6120
         TabIndex        =   19
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Buy"
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
      Begin vkUserContolsXP.vkTextBox vkTextBox7 
         Height          =   375
         Left            =   5040
         TabIndex        =   17
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
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
      Begin vkUserContolsXP.vkTextBox vkTextBox6 
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         Top             =   600
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
      Begin vkUserContolsXP.vkTextBox vkTextBox5 
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Top             =   600
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
      Begin vkUserContolsXP.vkTextBox vkTextBox4 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   600
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkTextBox vkTextBox3 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
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
      Begin VB.CommandButton Command1 
         Caption         =   "Þ"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   6000
         TabIndex        =   12
         Top             =   360
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ý"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6720
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   645
      End
      Begin vkUserContolsXP.vkLabel vkLabel7 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
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
      Begin vkUserContolsXP.vkLabel vkLabel6 
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "M.R.P."
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
         Left            =   4200
         TabIndex        =   8
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BackColor       =   16777215
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
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "Amount"
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
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "P.No"
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
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Select Type"
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2415
   End
   Begin vkUserContolsXP.vkListBox vkListBox2 
      Height          =   2895
      Left            =   4440
      TabIndex        =   2
      Top             =   1440
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5106
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0
   End
   Begin vkUserContolsXP.vkListBox vkListBox1 
      Height          =   6495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   11456
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0
   End
   Begin MSFlexGridLib.MSFlexGrid shopgrid 
      Height          =   2445
      Left            =   4440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5640
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   4313
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ForeColor       =   0
      Enabled         =   -1  'True
      SelectionMode   =   1
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

