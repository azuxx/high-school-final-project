<%@language="VBScript"%>
<html>
<%
set conn=server.createobject("ADODB.Connection")
conn.open "provider=microsoft.jet.oledb.4.0; data source=D:/Andrea/HTML/negozio_online/negozio_online.mdb"
set rs=server.createobject("ADODB.Recordset")
session("qta")=""
numart=session("artcar")
desc=session("descriz")
dd=session("stringa")
prez=session("prezunit")
For each item in request.form
		'response.write item
		if item<>"svuotaele" and item<>"controlla" then
			c = request.form(item)
			sql2="select a_id, a_disponibilità from Articoli where a_disponibilità>=" & c & " and a_id=" & item & ";"
			rs.open sql2, conn
			if not rs.eof then
				qta=qta & c
			else
				response.write "la quantità ordinata (" & c & " unità) è superiore alla disponibilità di prodotto" & item
				c=0
				qta=qta & c
			end if
			rs.close
		end if
next
for i=1 to len(qta)
	if mid(qta,i,1)<>"," and mid(qta,i,1)<>" " then
			dd1=dd1 & mid(qta,i,1)
	end if
next
session("qta")=dd1
dd1=session("qta")%>
<br>
<%'response.write dd1%>
<body background="images/litbg.gif">
<table border=1 background="images/fnd_pg_03.gif" style="font-family: Tahoma">
	<tr>
		<td align="center">Codice</td>
		<td align="center">Descrizione</td>
		<td align="center">prezzo unitario</td>
		<td align="center">Quantità</td>
		<td align="center" width="50">Totale</td>
	</tr>
	<%iniz=1
	iniz1=1
	for i=1 to numart%>
			<tr>
				<td><font face="Verdana" size="2"><%response.write request.cookies("carrello")("prodottoart" & i)%></td>
				<%for j=iniz to len(desc)
					if mid(desc,j,1)=";" then
						k=j%>
						<td><font face="Verdana" size="2"><%=mid(desc,iniz,k-iniz)%></td>
						<%iniz=k+1
						exit for
					end if
				next%>
				<%for y=iniz1 to len(prez)
					if mid(prez,y,1)=";" then
						k1=y%>
						<td width="100"><font face="Verdana" size="2"><%=formatcurrency(mid(prez,iniz1,k1-iniz1),2)%></td>
						<%pr=Cdbl(mid(prez,iniz1,k1-iniz1))%>
						<%iniz1=k1+1
						exit for
					end if
				next%>
				<td><p align="center"><%=mid(dd1,i,1)%></td>
				<%q=mid(dd1,i,1)
				sp=q*pr%>
				<td width="100"><%=formatcurrency(sp,2)%></td>
				<%sommatot=sommatot+sp%>
			</tr>
	<%next%>
	<tr>
	<td colspan=4>Importo totale carrello</td>
	<td width="100"><%=formatcurrency(sommatot,2)%></td>	
	</tr>
</table>
<br>
<%sql="select * from tpagamento"
rs.open sql, conn%>
<form name=frm6 action="ordine.asp" method=post>
Pagamento
<select size=1 name="pagscelta" id="pagscelta">
	<%while not rs.eof%> 
            <OPTION value="<%=rs("id_pagamento")%>"><%=rs("desc_pagamento")%></OPTION>
	<%rs.movenext	
	wend%>
</select>
<p>Indirizzo per la consegna <input type=text name="consegna" value=""></p>
<input type=hidden name="totale" value="<%=sommatot%>">
<input type=submit name="conferma" value="Conferma l'ordine">
</form>
<p>A seguito della vostra conferma Vi sarà inviata un e-mail contenente l'ordine 
da voi piazzato con le indicazioni dettagliate e il prezzo complessivo 
comprendente le spese di consegna e le eventuali spese per il pagamento. (Per 
ulteriori informazioni cliccare al menu-&gt; ASSISTENZA) </p>
<%rs.close
conn.close%>
</body>
</html>