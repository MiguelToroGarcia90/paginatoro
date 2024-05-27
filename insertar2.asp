<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,nom,apell,email,coments

nom=Request.Form("nombre")
apell=Request.Form("apellido")
email=Request.Form("correo")
coments=Request.Form("coments")

set conn=Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Miguel\Documents\PAGINA WRB\Compras1.mdb" 
conn.execute "INSERT INTO VENTAS(nombre,apellido,correo,coments) values('" & nom & "','" & coments & "','" & email &  "','" & coments & "')"
conn.close
set conn=nothing
response.redirect("VENTAS.html")

%>
</body>
</html>