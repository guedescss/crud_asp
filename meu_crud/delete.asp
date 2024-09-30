<!--#include file="conn.asp" -->
<%
Dim id, sql
id = Request.QueryString("id")

If id <> "" Then
    ' Usar JavaScript para confirmação
    Response.Write("<script>")
    Response.Write("if (confirm('Você tem certeza que deseja excluir este usuário?')) {")
    Response.Write("window.location = 'confirm_delete.asp?id=" & id & "';")
    Response.Write("} else {")
    Response.Write("window.location = 'index.asp';")
    Response.Write("}")
    Response.Write("</script>")
End If
%>
