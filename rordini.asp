<%@language=vbscript%>
<html>
<body background="images/fnd_pg_03.gif">
<%
id=session("idcliente")
if id<>"" then
	dim conn, rs, sql
	set conn=server.createobject("ADODB.connection")
	conn.open "provider=microsoft.jet.oledb.4.0; data source=D:/Andrea/HTML/negozio_online/negozio_online.mdb"
	set rs=server.createobject("ADODB.recordset")
	set rs2=server.createobject("ADODB.recordset")
	sql2="SELECT Dettagli_ord.dett_o_id, Articoli.a_descriz, Dettagli_ord.dett_qta FROM Articoli INNER JOIN Dettagli_ord ON Articoli.a_id = Dettagli_ord.dett_a_id ORDER BY Dettagli_ord.dett_o_id;"
	sql="SELECT Ordini.o_id, Ordini.o_data, Ordini.o_consegna, tpagamento.desc_pagamento, Ordini.o_totale FROM tpagamento INNER JOIN Ordini ON tpagamento.id_pagamento = Ordini.o_pagamento where o_c_id=" & id & ";"
	rs.open sql, conn
	rs2.open sql2, conn%>
	<table border=1>
		<%while not rs.eof%>
		<tr>
			<th style="border-top: 3px solid #000000">Codice ordine</th>
			<th style="border-top: 3px solid #000000">Data di effettuazione</th>
			<th style="border-top: 3px solid #000000">Luogo di consegna</th>
			<th style="border-top: 3px solid #000000">Pagamento</th>
			<th style="border-top: 3px solid #000000">Importo totale</th>
		</tr>
		<tr>
			<td><%=rs.fields("o_id")%></td>
			<td><%=rs.fields("o_data")%></td>
			<td><%=rs.fields("o_consegna")%></td>
			<td><%=rs.fields("desc_pagamento")%></td>
			<td><% response.write FormatCurrency(rs.fields("o_totale"),2)%></td>
		</tr>
		<tr>
			<th colspan=2>Prodotti acquistati</th>
		</tr>
			<%while not rs2.eof
				if rs("o_id")=rs2("dett_o_id") then%>
				<tr>
					<td><%=rs2.fields("a_descriz")%></td>
					<td><%=rs2.fields("dett_qta")%></td>
				</tr>
				<%end if
				rs2.movenext
			wend
			rs2.close
			rs2.open sql2, conn%>
		<%rs.movenext
		wend%>
	</table>
	<%rs.close
	conn.close%>
<%else%>
	<font size="4" face="Tahoma">Nessun cliente trovato!
Se non è stato effettuato alcun log in </font><a href=clienti.htm><img src=images/button_login.gif></a>
<%end if%>
</body>
</html>