<%@language="VBscript"%>
<html>
<body background="images/fondo.jpg">
<%
dim conn, rs, sql, login
set conn=server.createobject("ADODB.connection")
conn.open "provider=microsoft.jet.oledb.4.0; data source=D:/Andrea/HTML/negozio_online/negozio_online.mdb"
set rs=server.createobject("ADODB.recordset")
sql="Select * from Clienti where c_cognome='"& request.form("cognome") &"' and c_nome='"& request.form("nome") &"' and c_password='"& request.form("pass") &"';"
rs.open sql, conn
if rs.eof or session("login")=true then%><font color="#000080"> </font><font face="Tahoma" size="4">
<font color="#000080"><u>ACCESSO NEGATO!La password e/o l'username inseriti non sono corretti o è già attivo un profilo!</u></font> clicca per ritornare al log in</font> <a href=clienti.htm><img src=images/button_login.gif></a>
<%else
	id=rs("c_id")
	session("login")=true
	session("idcliente")=id
	session("nomelog")=request.form("nome")
	session("cognomelog")=request.form("cognome")
	id=session("idcliente")
	nomelog=session("nomelog")
	cognomelog=session("cognomelog")%>
	<font color="#000080"> </font><font face="Tahoma" size="4">
	<%response.write "Benvenuto " & nomelog & "  " & cognomelog & " ora puoi acquistare i nostri prodotti"
	'response.write id
end if
rs.close
conn.close%>
</body>
</html>