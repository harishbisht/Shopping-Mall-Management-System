VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim ind As Integer

Private Sub Form_Load()

Set rs = New ADODB.Recordset
sql = "select * from UserDetail"
'rs.Open sql, cn, adOpenDynamic, adLockOptimistic
rs.Open
rs.Update
'Combo1.List rs.Fields(UserName)
'Combo1.ItemData(0) = hhh
'For ind = 0 To rs.RecordCount - 1
rs.MoveFirst
Combo1.AddItem (rs.Fields(UserName))
rs.MoveNext
'Next ind
Combo1.AddItem "hitesh"
Call connection
End Sub
