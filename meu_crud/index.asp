<!--#include file="conn.asp" -->
<%
Dim rs, sql
sql = "SELECT * FROM Usuarios"
Set rs = conn.Execute(sql)
%>

<html>
<head>
    <title>CRUD Simples</title>
</head>
<body>
    <h1>Usuários</h1>
    <form action="add.asp" method="post">
        Nome: <input type="text" name="nome"><br>
        Email: <input type="text" name="email"><br>
        <input type="submit" value="Adicionar">
    </form>

    <h2>Lista de Usuários</h2>
    <table border="1">
        <tr>
            <th>ID</th>
            <th>Nome</th>
            <th>Email</th>
            <th>Ações</th>
        </tr>
        <% While Not rs.EOF %>
        <tr>
            <td><%= rs("ID") %></td>
            <td><%= rs("Nome") %></td>
            <td><%= rs("Email") %></td>
            <td>
                <a href="edit.asp?id=<%= rs("ID") %>">Editar</a>
                <a href="delete.asp?id=<%= rs("ID") %>">Excluir</a>
            </td>
        </tr>
        <% rs.MoveNext %>
        <% Wend %>
    </table>
</body>
</html>
<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>