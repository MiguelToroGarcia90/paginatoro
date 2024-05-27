<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,prod1,prod2,prod3,prod4,prod5,prod6,prod7,prod8,prod9,prod10,prod11,prod12,prod13,cant

prod1=Request.Form("prod1")
prod2=Request.Form("prod2")
prod3=Request.Form("prod3")
prod4=Request.Form("prod4")
prod5=Request.Form("prod5")
prod6=Request.Form("prod6")
prod7=Request.Form("prod7")
prod8=Request.Form("prod8")
prod9=Request.Form("prod9")
prod10=Request.Form("prod10")
prod11=Request.Form("prod11")
prod12=Request.Form("prod12")
cant=Request.Form("cantidad")

set conn=Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Miguel\Documents\PAGINA WRB\Compras1.mdb" 

if prod1=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar01', " & cint(cant) & ",45)"
end if 
if prod2=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar02', " & cint(cant) & ",20)"
end if 

if prod3=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar03', " & cint(cant) & ",11)"
end if 

if prod4=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar04', " & cint(cant) & ",11)"
end if 

if prod5=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar05', " & cint(cant) & ",44)"
end if 

if prod6=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar06', " & cint(cant) & ",23)"
end if 

if prod7=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar07', " & cint(cant) & ",23)"
end if 

if prod8=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar08', " & cint(cant) & ",21)"
end if 

if prod9=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar09', " & cint(cant) & ",15)"
end if 

if prod10=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar010', " & cint(cant) & ",45)"
end if 

if prod11=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar011', " & cint(cant) & ",35)"
end if 

if prod12=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar012', " & cint(cant) & ",21)"
end if 

if prod13=1 then
conn.execute "INSERT INTO Pedidos(Codigod,Cant,Precio) values( 'ar012', " & cint(cant) & ",45)"
end if 

conn.close
set conn=nothing
response.redirect("INDEX.html")

%>
</body>
</html>