Attribute VB_Name = "Module1"
Global con As ADODB.Connection
Global rs As ADODB.Recordset
'Public Function connectdb()
'Set con = New ADODB.Connection
' con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\crimefile.mdb;Persist Security Info=False")
'End Function
Public Sub dbconnection()
Set con = New ADODB.Connection
' Set res = New ADODB.Recordset
With con

   .ConnectionString = "Driver=(MySQL ODBC 5.1 Driver);SERVER=localhost;PWD=root;UID=root;PORT=3306;Data Source=crime"
   .CursorLocation = adUseClient
   .Open
  
End With
End Sub

