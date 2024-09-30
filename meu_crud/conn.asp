<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB;Data Source=10.100.2.11\SYSTEM30;Initial Catalog=exame;User ID=exame;Password=123456;"
%>

