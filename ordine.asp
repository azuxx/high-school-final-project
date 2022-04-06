<%@language="VBScript"%>
<html>
<body background="images/body_bg.jpg">
<%
numart=session("artcar")
dd1=session("qta")
id=session("idcliente")
'response.write id
data=date() & " " & time()
response.write data %>
<br>
<%set conn=server.createobject("ADODB.Connection")
conn.open "provider=microsoft.jet.oledb.4.0; data source=D:/Andrea/HTML/negozio_online/negozio_online.mdb"
set rs=server.createobject("ADODB.Recordset")
sql="insert into ordini(o_c_id,o_data,o_consegna,o_pagamento,o_totale) values(" & id & ",'" & data & "','" & request.form("consegna") & "','" & request.form("pagscelta") & "','" & request.form("totale") & "');"
rs.open sql, conn
sql2="SELECT TOP 1 Ordini.o_id FROM Ordini ORDER BY Ordini.o_id DESC;"
rs.open sql2, conn
if not rs.eof then
	ordid=rs("o_id")
end if
'response.write " " & ordid
rs.close
for i=1 to numart
	prod=request.cookies("carrello")("prodottoart" & i)
	quant=mid(dd1,i,1)
	sql3="insert into Dettagli_ord(dett_o_id,dett_a_id,dett_qta) values(" & ordid & "," & prod & "," & quant & ");"
	rs.open sql3, conn
next
sql4="UPDATE Ordini INNER JOIN (Articoli INNER JOIN Dettagli_ord ON Articoli.a_id=Dettagli_ord.dett_a_id) ON Ordini.o_id=Dettagli_ord.dett_o_id SET Articoli.a_disponibilità = (a_disponibilità)-(dett_qta) WHERE o_id="& ordid &";"
rs.open sql4, conn%><font face="BATAVIA" size="4"> L'ordine da lei richiesto è 
stato eseguito con successo!!!</font><p><font face="BATAVIA" size="4">Presto le 
sarà inviata un e-mail con la nostra conferma. La Zeta Elettronica la ringrazia 
di aver acquistato i suoi prodotti!</font></p>
<p><font face="BATAVIA" size="4"><a target="principale" href="home.htm">
<font color="#3399FF">Torna all'home</font></a></font></p>
<%session.Abandon%>
</body>
</html>