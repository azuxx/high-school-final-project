<%@language="VBScript"%>
<html>
<%
if session("login")=true then
		Session.Abandon 
		nomelog=session("nomelog")
		cognomelog=session("cognomelog")
		response.write "Si è disconnesso: " & nomelog & " " & cognomelog
else
	if session("login")=false then%>
		<u><font face="Tahoma" size="4" color="#000080">Impossibile effetuare il log out perchè non è avvenuto alcun accesso!!</font></u>
	<%end if
end if%>
<body background="images/litbg.gif">
</body>
</html>