<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,nom,email,coments

nom=Request.Form("usuario")
email=Request.Form("correo")
coments=Request.Form("coments")

set conn=Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Miguel\Documents\PAGINA WRB\Compras1.mdb" 
conn.execute "INSERT INTO comentarios(usuario,correo,coments) values('" & nom & "','" & email &  "','" & coments & "')"
conn.close
set conn=nothing
response.redirect("comentarios.html")

%>
</body>
</html>