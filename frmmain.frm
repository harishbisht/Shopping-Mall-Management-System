VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15240
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Left            =   8040
      Top             =   7320
   End
   Begin VB.Timer Timer3 
      Left            =   7320
      Top             =   7320
   End
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   4800
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      HasDC           =   0   'False
      Height          =   9915
      Left            =   0
      ScaleHeight     =   9855
      ScaleWidth      =   15180
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin VB.PictureBox Picture2 
         Height          =   1695
         Left            =   4800
         Picture         =   "frmmain.frx":0000
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         Height          =   1575
         Left            =   4800
         Picture         =   "frmmain.frx":09D6
         ScaleHeight     =   1515
         ScaleWidth      =   1635
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin vkUserContolsXP.vkFrame frame1 
         Height          =   10575
         Left            =   2040
         TabIndex        =   4
         Top             =   0
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   18653
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
         BreakCorner     =   0   'False
         Begin ButtonXp.XPButton XPButton3 
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Command3"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
      End
      Begin VB.Timer Timer1 
         Left            =   840
         Top             =   4680
      End
      Begin vkUserContolsXP.vkFrame framemain 
         Height          =   8535
         Left            =   -240
         TabIndex        =   1
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   15055
         Caption         =   "Menu"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin ButtonXp.XPButton XPButton5 
            Height          =   375
            Left            =   600
            TabIndex        =   8
            Top             =   3960
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
            Caption         =   "Command5"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin ButtonXp.XPButton XPButton4 
            Height          =   495
            Left            =   600
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
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
            Caption         =   "New User"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin ButtonXp.XPButton XPButton2 
            Height          =   495
            Left            =   600
            TabIndex        =   5
            Top             =   480
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
            Caption         =   "New User"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin ButtonXp.XPButton XPButton1 
            Height          =   375
            Left            =   480
            TabIndex        =   2
            Top             =   7440
            Width           =   1215
            _ExtentX        =   2143
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
            Caption         =   "Close"
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   0
            Width           =   2055
         End
      End
   End
   Begin VB.Menu mnuuser 
      Caption         =   "&User"
      Begin VB.Menu mnulogout 
         Caption         =   "&Log out"
      End
      Begin VB.Menu mnuchangepass 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnudash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&exit"
      End
   End
   Begin VB.Menu mnuitem 
      Caption         =   "&Item"
      Begin VB.Menu mnuadd 
         Caption         =   "Add Item"
      End
      Begin VB.Menu mnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnubutitem 
         Caption         =   "&Buy Item"
      End
   End
   Begin VB.Menu mnucustomer 
      Caption         =   "&Customer"
      Begin VB.Menu mnuaddcustomer 
         Caption         =   "C&ustomer Details"
      End
   End
   Begin VB.Menu mnudist 
      Caption         =   "&Distributer"
      Begin VB.Menu mnudistdetails 
         Caption         =   "Distributer  De&tail"
      End
      Begin VB.Menu mnupurchaseitem 
         Caption         =   "&Purchase Item"
      End
   End
   Begin VB.Menu mnustock 
      Caption         =   "&Stock"
      Begin VB.Menu mnulist 
         Caption         =   "&Product List"
      End
      Begin VB.Menu mnuexpired 
         Caption         =   "&Expired Products"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report"
      Begin VB.Menu mnusellingreport 
         Caption         =   "&Selling Report"
      End
      Begin VB.Menu mnudashr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnupurchasereport 
         Caption         =   "&Purchase report"
      End
      Begin VB.Menu mnudashr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnucustomerreport 
         Caption         =   "&Customer Report"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
      Begin VB.Menu mnucontact 
         Caption         =   "&Contact Us"
      End
      Begin VB.Menu mnuaboutus 
         Caption         =   "&About Us"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msg
Dim click

'Private Sub cmdnewuser_Click()
'Form3.Show
'Unload Me
'End Sub

Private Sub Label1_Click()
framemain.Height = 1
Timer1.Interval = 1
Timer1.Enabled = True

If framemain.Height >= 10575 Then
framemain.Height = 10575
Timer1.Enabled = False
framemain.Height = 1
Timer1.Interval = 1
Timer1.Enabled = True

End If
End Sub

Private Sub MDIForm_Load()
framemain.Height = 1
Timer1.Interval = 1
Timer1.Enabled = True
Label1.Enabled = False
End Sub

Private Sub mnuadd_Click()
Form9.Show
End Sub

Private Sub mnuaddcustomer_Click()
Form5.Show
Unload Me
End Sub

Private Sub mnubutitem_Click()
Form7.Show
End Sub

Private Sub mnuchangepass_Click()
Form2.Show
'Unload Me
End Sub

Private Sub mnudistdetails_Click()
Form6.Show
End Sub

Private Sub mnuexit_Click()
msg = MsgBox("Do you want to quit?", vbYesNo + vbQuestion)
If msg = vbYes Then
End
'Else
'Load MDIForm1
End If
End Sub

Private Sub mnulogout_Click()
click = MsgBox("Do u want to quit?", vbYesNo + vbQuestion)
If click = vbYes Then
End
End If
End Sub

Private Sub mnupurchaseitem_Click()
Form8.Show
Unload Me
End Sub

Private Sub mnureUserDetail_Click()
msg = MsgBox(" You are UserDetailg off. Do you want to continue", vbYesNo + vbQuestion)
If msg = vbYes Then
Unload Me
Form1.Show
Else
End If
End Sub

Private Sub mnurelogin_Click()
Form1.Show

End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = True
Picture3.Visible = False
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = True
Picture2.Visible = False
End Sub

Private Sub Timer1_Timer()
framemain.Height = framemain.Height + 200
If framemain.Height >= 10575 Then
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
framemain.Height = framemain.Height - 100
If framemain.Height <= 300 Then
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
frame1.Width = frame1.Width + 200
If frame1.Width >= 10575 Then
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
frame1.Width = frame1.Width - 200
If frame1.Width <= 100 Then
Timer4.Enabled = False
End If
End Sub

Private Sub XPButton1_Click()
framemain.Height = 10575
Timer2.Interval = 1
Timer2.Enabled = True
Label1.Enabled = True
End Sub

Private Sub XPButton2_Click()
XPButton2.Visible = False
XPButton4.Visible = True
frame1.Width = 1
Timer3.Interval = 1
Timer3.Enabled = True
'XPButton2.Visible = False
'XPButton4.Visible = True
End Sub

Private Sub XPButton3_Click()
XPButton2.Visible = True
XPButton4.Visible = False
frame1.Width = 8000
Timer4.Interval = 1
Timer4.Enabled = True
End Sub

Private Sub XPButton4_Click()
XPButton2.Visible = True
XPButton4.Visible = False
frame1.Width = 8700
Timer4.Interval = 1
Timer4.Enabled = True
End Sub

