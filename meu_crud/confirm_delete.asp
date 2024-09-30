<!--#include file="conn.asp" -->
<%
Dim id, sql
id = Request.QueryString("id")

If id <> "" Then
    sql = "DELETE FROM Usuarios WHERE ID=" & id
    conn.Execute(sql)
End If

Response.Redirect "index.asp"
%>
