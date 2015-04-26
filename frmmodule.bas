Attribute VB_Name = "Module1"
Public cn As New ADODB.connection

Sub main()
Form1.Show
End Sub

Public Sub connection()
If cn.State = 1 Then
cn.Close
End If
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\sms\hr.mdb;Persist Security Info=False"
End Sub
