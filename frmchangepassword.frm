VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Form2 
   Caption         =   "Change Password"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form2"
   ScaleHeight     =   7740
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   5535
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9763
      Caption         =   "Change Password"
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
         Left            =   2040
         TabIndex        =   12
         Top             =   4320
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
         Caption         =   "&Clear"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdok 
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   4320
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
      Begin ButtonXp.XPButton cmdcancel 
         Height          =   495
         Left            =   3360
         TabIndex        =   10
         Top             =   4320
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
         Caption         =   "Ca&ncel"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Confirm Password"
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
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "New Password"
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
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Old Password"
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
      Begin vkUserContolsXP.vkTextBox txtconfpass 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   3120
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
      Begin vkUserContolsXP.vkTextBox txtnewpass 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   2400
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
      Begin vkUserContolsXP.vkTextBox txtoldpass 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   1680
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
      Begin vkUserContolsXP.vkLabel vkLabel5 
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Usename"
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
      Begin vkUserContolsXP.vkTextBox txtusername 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
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
   End
   Begin vkUserContolsXP.vkLabel vkLabel4 
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Change Password"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String
Dim flag As Boolean
Dim chpuser As String

Private Sub cmdcancel_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub cmdclear_Click()
txtusername.Text = ""
txtconfpass.Text = ""
txtnewpass.Text = ""
txtoldpass.Text = ""
End Sub

Private Sub cmdok_Click()
If txtusername.Text = "" Or txtconfpass.Text = "" Or txtnewpass.Text = "" Or txtoldpass.Text = "" Then
MsgBox "Please fill all the fields"
Exit Sub
End If
'Exit Sub
Set rs = New ADODB.Recordset
chpuser = txtusername.Text
sql = "select * from UserDetail where Username='" & chpuser & "'"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic

If Len(txtnewpass.Text) < 6 Then
MsgBox "Password should not less than 6 character"
txtnewpass.Text = ""
txtconfpass.Text = ""
Exit Sub
'txtnewpass.Text = SetFocus
ElseIf txtnewpass.Text <> txtconfpass.Text Then
MsgBox ("Password do not match!")
txtnewpass.Text = ""
txtconfpass.Text = ""
Exit Sub
End If
If rs.EOF = False Then
    If rs.Fields("Password") = txtoldpass.Text Then
        rs.Fields("Password") = txtnewpass.Text
        rs.Update
    Else
        MsgBox ("Old Password do not match!")
        txtoldpass.Text = ""
        Exit Sub
    End If
'rs.Update
MsgBox ("Password successfully change!!!")
'rs.Close
Me.Hide
MDIForm1.Show
End If

End Sub

Private Sub Form_Load()
Call connection
Set rs = New ADODB.Recordset
sql = "select * from UserDetail"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
End Sub


