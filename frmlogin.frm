VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin vkUserContolsXP.vkLabel time 
      Height          =   735
      Left            =   1800
      TabIndex        =   11
      Top             =   3720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   16777215
      Caption         =   "30"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5520
      Top             =   3360
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5040
      Top             =   2160
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
   Begin VB.ComboBox cmbusertype 
      Height          =   315
      ItemData        =   "frmlogin.frx":0000
      Left            =   1560
      List            =   "frmlogin.frx":000A
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   4815
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8493
      Caption         =   "Login"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ProgressBar sms 
         Height          =   2535
         Left            =   4560
         TabIndex        =   0
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   4471
         _Version        =   393216
         Appearance      =   0
         Max             =   30
         Orientation     =   1
         Scrolling       =   1
      End
      Begin vkUserContolsXP.vkToggleButton cmdforgotpass 
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   3360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "Forgot Your Password"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkTextBox txtpassword 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   2520
         Width           =   2415
         _ExtentX        =   4260
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
         PassWordChar    =   "*"
         LegendForeColor =   12563634
      End
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "User Type"
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
         Left            =   480
         TabIndex        =   9
         Top             =   2640
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Password"
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
         Left            =   480
         TabIndex        =   8
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Username"
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
      Begin ButtonXp.XPButton cmdcancel 
         Height          =   495
         Left            =   2880
         TabIndex        =   6
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
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
      Begin ButtonXp.XPButton cmdok 
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "&Ok"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin vkUserContolsXP.vkTextBox txtusername 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
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
         LegendForeColor =   12563634
      End
      Begin VB.Shape Shape3 
         Height          =   375
         Left            =   -360
         Top             =   -480
         Width           =   5175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim count1 As Integer
Dim sql As String
Dim click

Private Sub cmdCancel_Click()
click = MsgBox("Do u want to quit?", vbYesNo + vbQuestion)
If click = vbYes Then
End
'Else
'Load Form1
End If
End Sub


Private Sub cmdforgotpass_Click()
Form4.Show
Me.Hide
End Sub

Private Sub cmdok_Click()
Dim flag As Boolean
flag = False

Set rs = New ADODB.Recordset
sql = "select * from UserDetail"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
'rs.MoveLast

If cmbusertype.Text = "" Then
MsgBox "Please Select the User Type"

ElseIf (txtusername.Text = "" And txtpassword.Text = "") Then
    MsgBox "Enter username and password!!"
Else
    Do While Not rs.EOF
    rs.MoveLast
If rs.Fields(0) = cmbusertype.Text And rs.Fields(1) = txtusername.Text And rs.Fields(2) = txtpassword.Text Then
   flag = True
   Exit Do
End If
        rs.MoveNext
    Loop
        If flag = True Then
            MDIForm1.Show
            Unload Me
        
        Else
            MsgBox ("Your Username or Password is incorrect!!")
            txtusername.Text = ""
            txtpassword.Text = ""
            count1 = count1 + 1
        If (count1 = 3) Then
         End
        End If
    End If
End If
End Sub

Private Sub Form_Load()

Call connection
Set rs = New ADODB.Recordset
sql = "select * from UserDetail"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
'rs.MoveFirst
count1 = 0
End Sub

Private Sub Timer1_Timer()
If sms.Value > 30 Then
MsgBox "Sorry!! Time-Out"
End
End If

sms.Value = sms.Value + 0.015
time.Caption = sms.Value
If time.Caption = 30 Then
   time.Caption = 1
ElseIf time.Caption = 29 Then
   time.Caption = 2
ElseIf time.Caption = 28 Then
   time.Caption = 3
ElseIf time.Caption = 27 Then
   time.Caption = 4
ElseIf time.Caption = 26 Then
   time.Caption = 5
ElseIf time.Caption = 25 Then
   time.Caption = 6
ElseIf time.Caption = 24 Then
   time.Caption = 7
ElseIf time.Caption = 23 Then
   time.Caption = 8
ElseIf time.Caption = 22 Then
   time.Caption = 9
ElseIf time.Caption = 21 Then
   time.Caption = 10
ElseIf time.Caption = 20 Then
   time.Caption = 11
ElseIf time.Caption = 19 Then
   time.Caption = 12
ElseIf time.Caption = 18 Then
   time.Caption = 13
ElseIf time.Caption = 17 Then
   time.Caption = 14
ElseIf time.Caption = 16 Then
   time.Caption = 15
ElseIf time.Caption = 15 Then
   time.Caption = 16
ElseIf time.Caption = 14 Then
   time.Caption = 17
ElseIf time.Caption = 13 Then
   time.Caption = 18
ElseIf time.Caption = 12 Then
   time.Caption = 19
ElseIf time.Caption = 11 Then
   time.Caption = 20
ElseIf time.Caption = 10 Then
   time.Caption = 21
ElseIf time.Caption = 9 Then
   time.Caption = 22
ElseIf time.Caption = 8 Then
   time.Caption = 23
ElseIf time.Caption = 7 Then
   time.Caption = 24
ElseIf time.Caption = 6 Then
   time.Caption = 25
ElseIf time.Caption = 5 Then
   time.Caption = 26
ElseIf time.Caption = 4 Then
   time.Caption = 27
ElseIf time.Caption = 3 Then
   time.Caption = 28
ElseIf time.Caption = 2 Then
   time.Caption = 29
ElseIf time.Caption = 1 Then
   time.Caption = 30
End If
bye:
Exit Sub
End Sub
