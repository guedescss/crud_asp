<!--#include file="conn.asp" -->
<%
Dim id, nome, email, sql, rs
id = Request.QueryString("id")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    nome = Request.Form("nome")
    email = Request.Form("email")
    sql = "UPDATE Usuarios SET Nome='" & nome & "', Email='" & email & "' WHERE ID=" & id
    conn.Execute(sql)
    Response.Redirect "index.asp"
Else
    sql = "SELECT * FROM Usuarios WHERE ID=" & id
    Set rs = conn.Execute(sql)
    nome = rs("Nome")
    email = rs("Email")
    rs.Close
End If
%>

<html>
<head>
    <title>Editar Usuário</title>
</head>
<body>
    <h1>Editar Usuário</h1>
    <form action="edit.asp?id=<%= id %>" method="post">
        Nome: <input type="text" name="nome" value="<%= nome %>"><br>
        Email: <input type="text" name="email" value="<%= email %>"><br>
        <input type="submit" value="Salvar">
    </form>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>
