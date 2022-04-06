<%@language="VBScript"%>
<html>
<head>
<script language="vbscript">
sub carrello_onclick(j)
	frm3.cscelta.value=j
	frm3.submit
end sub
</script>
</head>
<body background="images/litbg.gif">
<%dim sql, sql2, conn, rs
if request.form("svuotaele")<>"" then
	Session("artcar") = 0
	session("stringa")=""
	session("qta")=""
	session("descriz")=""
	session("prezunit")=""
end if
dd=session("stringa")
'response.write "gli art sono" & dd
if request.form("scelta")<>"" then

	set conn=server.createobject("ADODB.Connection")
	conn.open "provider=microsoft.jet.oledb.4.0; data source=D:/Andrea/HTML/negozio_online/negozio_online.mdb"
	set rs=server.createobject("ADODB.Recordset")%>
	<form name=frm3 action=carrello.asp method=post>
	<%sql="select * from categorie where cat_id=" & request.form("scelta") & ";"
	rs.open sql, conn%>
	<br>
	<font face="Tahoma" size=5><%=rs.fields("cat_descriz")%></font>
	<p></p>
	<%rs.close
	sql2="Select * from Articoli where a_cat_id=" & request.form("scelta") & ";"
	rs.open sql2, conn%>
	<table border=1>
		<tr>
			<td align="center">Codice</td>
			<td align="center">Foto</td>
			<td align="center">Descrizione</td>
			<td align="center">prezzo unitario<br>(in <font face="Times New Roman">
			€)</font></td>
			<td align="center">Disponibilità</td>
			<td align="center">Quantità</td>
		</tr>
	<% while not rs.EOF%>
		<tr>
			<td><%=rs.fields("a_id")%></td>
			<%if rs.fields("a_foto")<>"" then%>
				<td>
				<p align="center"><img src="images/<%=rs.fields("a_foto")%>"></td>
			<%else%>
				<td>anteprima non disponibile</td>
			<%end if%>
			<td><%=rs.fields("a_descriz")%></td>
			<td align="right"><%=formatcurrency(rs.fields("a_prezunit"),2)%></td>
			<td align="center"><%=rs.fields("a_disponibilità")%></td><%quantit=rs("a_id")%>
			<td>
			<input type=text name="<%=quantit%>" value="" maxlength=1 size="8"></td>
			<%
			s=0
			iniz2=1
			for i=iniz2 to len(dd)
				if mid(dd,i,1)=";" then
					k=i
					codart=mid(dd,iniz2,k-iniz2)
					iniz2=k+1
					if codart=Cstr(rs("a_id")) then
						s=s+1
					end if
				end if
			next
			if session("login")=true and s=0 and rs.fields("a_disponibilità")<>0 then%>
						<td>
						<img name=carrello src="images/icone_proc_me_elettr.jpg" onclick='carrello_onclick(<%=rs.fields("a_id")%>)' width="84" height="73"></td>
			<%end if%>
		</tr>
		<%rs.movenext
	wend%>
	</table>
	<input type=hidden name=cscelta id=cscelta value="">
	<%rs.close
	conn.close%>
	</form>
<%else%>
	<font face="Tahoma" size="4">Carrello svuotato! torna al catalogo...
</font>
<%end if%>
</body>
</html>