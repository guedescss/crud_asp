<!--#include file="conn.asp" -->
<%
Dim nome, email, sql
nome = Request.Form("nome")
email = Request.Form("email")

If nome <> "" And email <> "" Then
    sql = "INSERT INTO Usuarios (Nome, Email) VALUES ('" & nome & "', '" & email & "')"
    conn.Execute(sql)
End If

Response.Redirect "index.asp"
%>