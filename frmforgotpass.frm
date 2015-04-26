VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Form4 
   Caption         =   "Forgot Password"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   2910
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin vkUserContolsXP.vkFrame frameforgotpass 
      Height          =   2895
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5106
      Caption         =   "Forgot Password"
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
         Left            =   2040
         TabIndex        =   9
         Top             =   2040
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
         Left            =   600
         TabIndex        =   8
         Top             =   2040
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
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Mobile No."
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
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
      Begin vkUserContolsXP.vkTextBox txtmobileno 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
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
         MaxLength       =   10
         LegendAlignmentX=   0
         LegendForeColor =   12563634
      End
      Begin vkUserContolsXP.vkTextBox txtusername 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   720
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
         LegendForeColor =   12563634
      End
   End
   Begin vkUserContolsXP.vkFrame framepass 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5106
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
      Begin ButtonXp.XPButton cmdokk 
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   2040
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
      Begin vkUserContolsXP.vkTextBox txtforgotpass 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         ForeColor       =   255
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         LegendForeColor =   12563634
      End
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Yout Password is shown below "
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
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String
Private Sub cmdcancel_Click()
Form1.Show
Me.Hide
End Sub

Private Sub cmdclear_Click()
txtusername.Text = ""
txtmobileno.Text = ""
End Sub

Private Sub cmdok_Click()
If txtmobileno.Text = "" And txtusername.Text <> "" Then
    MsgBox "ENter Yout Mobile Number"
ElseIf txtmobileno.Text <> "" And txtusername.Text = "" Then
    MsgBox "Enter Your Username"
ElseIf txtmobileno.Text = "" And txtusername.Text = "" Then
    MsgBox "Fill all Field"
    txtusername.SetFocus
ElseIf Len(txtmobileno.Text) < 10 Then
MsgBox "Please Enter a valid Mobile number"
txtmobileno.Text = ""
Else
Set rs = New ADODB.Recordset
fpuser = txtusername.Text
sql = "select * from UserDetail where Username='" & fpuser & "'"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic

If rs.EOF = False Then
If rs.Fields("Mobile") = txtmobileno.Text And rs.Fields("Username") = txtusername.Text Then
framepass.Visible = True
frameforgotpass.Visible = False

txtforgotpass.Text = rs.Fields("Password")
rs.Update
rs.Close
'Unload Me
End If
End If
End If
End Sub

Private Sub cmdokk_Click()
txtforgotpass.Text = ""
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
frameforgotpass.Visible = True
framepass.Visible = False

Call connection
Set rs = New ADODB.Recordset
sql = "select * from UserDetail"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
End Sub


Private Sub txtmobileno_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
    If KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
    End If
End Sub
