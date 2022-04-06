<%@language="VBScript"%>
<html>
<%dim conn, rs, sql
set conn=server.createobject("ADODB.connection")
conn.open "provider=microsoft.jet.oledb.4.0; data source=D:/Andrea/HTML/negozio_online/negozio_online.mdb"
set rs=server.createobject("ADODB.recordset")
sql="insert into Clienti(c_cognome,c_nome,c_datanas,c_indirizzo,c_città,c_email,c_password) values('" & request.form("rcognome") & "','" & request.form("rnome") & "','" & request.form("rdatanas") & "','" & request.form("rindirizzo") & "','" & request.form("rcitta") & "','" & request.form("rmail")& "','" & request.form("rpass") &"');" 
rs.open sql, conn%>
<body background="images/fondo.jpg">

<i><font face="High Tower Text" size="5">REGISTRAZIONE EFFETTUATA CON SUCCESSO!
</font></i>
<p><i><font face="High Tower Text" size="5">puoi stampare la pagina con i tuoi dati personali
</font></i></p>
<table border=1>
<tr>
	<td><%response.write request.form("rnome")%></td>
</tr>
<tr>
	<td><%response.write request.form("rcognome")%></td>
</tr>
<tr>
	<td><%response.write request.form("rdatanas")%></td>
</tr>
<tr>
	<td><%response.write request.form("rindirizzo")%></td>
</tr>
<tr>
	<td><%response.write request.form("rcitta")%></td>
</tr>
<tr>
	<td><%response.write request.form("rmail")%></td>
</tr>
<tr>
	<td><%response.write request.form("rpass")%></td>
</tr>
<tr>
	<td><%response.write "E' stata accettato il seguente contratto di registrazione"%></td>
</tr>
</table>
</html>