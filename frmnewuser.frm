VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "New Login"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form3"
   ScaleHeight     =   8430
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin vkUserContolsXP.vkFrame frameregestration 
      Height          =   8295
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   14631
      Caption         =   "Registration Form"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkLabel vkLabel11 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "UserType"
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
      Begin VB.ComboBox cmbusertype 
         Height          =   315
         ItemData        =   "frmnewuser.frx":0000
         Left            =   1680
         List            =   "frmnewuser.frx":000A
         TabIndex        =   24
         Top             =   840
         Width           =   2415
      End
      Begin vkUserContolsXP.vkLabel vkLabel10 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   855
         _ExtentX        =   1508
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
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3720
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
      Begin vkUserContolsXP.vkTextBox txtusername 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   3600
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
      Begin ButtonXp.XPButton cmdclear 
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Top             =   7320
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
         Caption         =   "&Clear All"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin ButtonXp.XPButton cmdok 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   7320
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
         Left            =   2880
         TabIndex        =   12
         Top             =   7320
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
      Begin vkUserContolsXP.vkTextBox txtfirstname 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
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
         LegendForeColor =   12563634
      End
      Begin vkUserContolsXP.vkLabel vkLabel9 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Last Name"
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
         Left            =   120
         TabIndex        =   20
         Top             =   4320
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
      Begin vkUserContolsXP.vkLabel vkLabel7 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   4920
         Width           =   1335
         _ExtentX        =   2355
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
      Begin vkUserContolsXP.vkLabel vkLabel6 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   5520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "City"
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
         Left            =   120
         TabIndex        =   17
         Top             =   6720
         Width           =   855
         _ExtentX        =   1508
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
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   6120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Pin Code"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "First Name"
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
      Begin vkUserContolsXP.vkTextBox txtlastname 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   2280
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
      Begin vkUserContolsXP.vkTextBox txtpass 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   4200
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
      Begin vkUserContolsXP.vkTextBox txtconfpass 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   4800
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
      Begin vkUserContolsXP.vkTextBox txtcity 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   5400
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
      Begin vkUserContolsXP.vkTextBox txtpincode 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   6000
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
         MaxLength       =   6
         LegendForeColor =   12563634
      End
      Begin vkUserContolsXP.vkTextBox txtmobileno 
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   6600
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
         MaxLength       =   10
         LegendForeColor =   12563634
      End
      Begin vkUserContolsXP.vkOptionButton optmale 
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Male"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton optfemale 
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Female"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4920
      Top             =   6000
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\STUDY\Projects\Hitesh Proj\hr.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\STUDY\Projects\Hitesh Proj\hr.mdb;Persist Security Info=False"
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
   Begin vkUserContolsXP.vkLabel vkLabel2 
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Registration Form"
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql As String

Private Sub cmdCancel_Click()
Form1.Show
Me.Hide
End Sub

Private Sub cmdclear_Click()
txtcity.Text = ""
txtconfpass.Text = ""
txtfirstname.Text = ""
txtlastname.Text = ""
txtmobileno.Text = ""
txtpass.Text = ""
txtpincode.Text = ""
txtUserName.Text = ""

End Sub

Private Sub cmdOK_Click()
If cmbusertype.Text = "" Or txtpass.Text = "" Or txtconfpass.Text = "" Or txtfirstname.Text = "" Or txtlastname.Text = "" Or txtmobileno.Text = "" Or txtpass.Text = "" Or txtpincode.Text = "" Or txtUserName.Text = "" Then
MsgBox "Fill all the field"
ElseIf Len(txtpass.Text) < 6 Or Len(txtconfpass.Text) < 6 Then
MsgBox "Password should not less than 6 character"
txtpass.Text = ""
txtpass.SetFocus
ElseIf (txtpass.Text <> txtconfpass.Text) Then
MsgBox "Re-enter  the Password"
txtpass.Text = ""
txtconfpass.Text = ""
txtpass.SetFocus
Else
rs.AddNew
If optmale.Value = True Or optfemale.Value = False Then
rs.Fields("Gender") = "Male"
ElseIf optfemale.Value = True Or optmale.Value = False Then
rs.Fields("Gender") = "Female"
End If
 rs.Fields("Usertype") = cmbusertype.Text
 rs.Fields("Username") = txtUserName.Text
 rs.Fields("Password") = txtpass.Text
 rs.Fields("Firstname") = txtfirstname.Text
 rs.Fields("Lastname") = txtlastname.Text
 rs.Fields("Pincode") = txtpincode.Text
 rs.Fields("City") = txtcity.Text
 rs.Fields("Mobile") = txtmobileno.Text

 rs.Update
MsgBox "New User added successfully"
txtcity.Text = ""
txtconfpass.Text = ""
txtfirstname.Text = ""
txtlastname.Text = ""
txtmobileno.Text = ""
txtpass.Text = ""
txtpincode.Text = ""
txtUserName.Text = ""
Form1.Show
Me.Hide
End If
End Sub

Private Sub Form_Load()
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


Private Sub txtpincode_KeyPress(KeyAscii As Integer)

If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
    If KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
    End If
End Sub
