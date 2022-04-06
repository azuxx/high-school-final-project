<%@language="VBscript"%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">

<head>
<!--[if !mso]>
<style>
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
</style>
<![endif]-->
<link rel="File-List" href="carrello_file/filelist.xml">
<script language="vbscript">
sub controlla_onclick()
	frm5.action="controlla.asp"
end sub
sub svuota_onclick()
	frm5.svuotaele.value="svuotato"
	frm5.action="articoli.asp"
end sub
</script>
<!--[if gte mso 9]>
<xml><o:shapedefaults v:ext="edit" spidmax="1027"/>
</xml><![endif]-->
</head>

<body background="images/sfondo.gif">
<p>&nbsp;</p>
<p><!--[if gte vml 1]><v:shapetype id="_x0000_t136"
 coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">
 <v:formulas>
  <v:f eqn="sum #0 0 10800"/>
  <v:f eqn="prod #0 2 1"/>
  <v:f eqn="sum 21600 0 @1"/>
  <v:f eqn="sum 0 0 @2"/>
  <v:f eqn="sum 21600 0 @3"/>
  <v:f eqn="if @0 @3 0"/>
  <v:f eqn="if @0 21600 @1"/>
  <v:f eqn="if @0 0 @2"/>
  <v:f eqn="if @0 @4 21600"/>
  <v:f eqn="mid @5 @6"/>
  <v:f eqn="mid @8 @5"/>
  <v:f eqn="mid @7 @8"/>
  <v:f eqn="mid @6 @7"/>
  <v:f eqn="sum @6 0 @5"/>
 </v:formulas>
 <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"
  o:connectangles="270,180,90,0"/>
 <v:textpath on="t" fitshape="t"/>
 <v:handles>
  <v:h position="#0,bottomRight" xrange="6629,14971"/>
 </v:handles>
 <o:lock v:ext="edit" text="t" shapetype="t"/>
</v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t136" alt="Carrello"
 style='width:198.75pt;height:33pt' strokecolor="#eaeaea" strokeweight=".35189mm">
 <v:fill src="carrello_file/image001.jpg" o:title="contact" color2="blue"
  recolor="t" colors="0 #a603ab;13763f #0819fb;22938f #1a8d48;34079f yellow;47841f #ee3f17;57672f #e81766;1 #a603ab"
  method="none" type="frame"/>
 <v:shadow on="t" type="perspective" color="silver" opacity="52429f" origin="-.5,.5"
  matrix=",46340f,,.5,,-4768371582e-16"/>
 <v:textpath style='font-family:"Cracked Johnnie";font-size:32pt;v-text-kern:t'
  trim="t" fitpath="t" string="Carrello"/>
</v:shape><![endif]--><![if !vml]><img border=0 width=282 height=48
src="carrello_file/image002.gif" alt=Carrello v:shapes="_x0000_s1025"><![endif]></p>
<%
Dim prodotto, numart
prodotto = Request.Form("cscelta")
if prodotto<>"" then
	desc=session("descriz")
	prez=session("prezunit")
	set conn=server.createobject("ADODB.Connection")
	conn.open "provider=microsoft.jet.oledb.4.0; data source=D:/Andrea/HTML/negozio_online/negozio_online.mdb"
	set rs=server.createobject("ADODB.Recordset")
	sql="Select * from Articoli where a_id=" & request.form("cscelta") & ";"
	rs.open sql, conn
	while not rs.eof
		desc=desc & rs.fields("a_descriz") & ";"
		prez=prez & rs.fields("a_prezunit") & ";"
		session("descriz")=desc
		session("prezunit")=prez
		rs.movenext
	wend
	rs.close
end if
numart=session("artcar")
desc=session("descriz")
prez=session("prezunit")
dd=session("stringa")
dd1=session("qta")
if session("login")=true then
	if prodotto<>"" then
		numart=numart+1
		response.cookies("carrello")("prodottoart" & numart)=prodotto
	end if
	if numart<>0 then%>
	<form name=frm5 action="" method=post>
	<table border=1 background="images/bgd.gif" style="font-family: Tahoma">
			<tr>
				<td align="center">numero elemento</td>
				<td align="center">codice articolo</td>
				<td align="center">descrizione</td>
				<td align="center">prezzo unitario</td>
				<td align="center">quantità ordinata (modificabile da 1 a 9)</td>
			</tr>
		<%iniz=1
		iniz1=1
		iniz2=1%>
		<%for i=1 to numart%>
		<tr>
			<td><font face="Verdana" size="2"><%response.write i%></td>
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
					<td><font face="Verdana" size="2"><%=formatcurrency(mid(prez,iniz1,k1-iniz1),2)%></td>
					<%iniz1=k1+1
					exit for
				end if
			next%>
			<%For each item in request.form
					if item <> "cscelta" and item=request.cookies("carrello")("prodottoart" & i) then
							c = request.form(item)
							'response.write(c & " - " & item & ",")
							if c <> "" and c <> " " then
								sql2="select a_id, a_disponibilità from Articoli where a_disponibilità>=" & c & " and a_id=" & item & ";"
								rs.open sql2, conn
								if not rs.eof then
									dd = dd & item & ";"
									dd1=dd1 & c
								else%>
									<br>
									<%dd = dd & item & ";"
									c=0
									dd1=dd1 & c%>
									<font size="4">Attenzione la quantità del prodotto ordinato è superiore alla sua disponibilità in magazzino! La invitiamo ad inserire una quantità inferiore.</font>
									<br>
								<%end if
								rs.close
							end if
					end if
				next%>
			<td>
			<%for x=iniz2 to len(dd)
				if mid(dd,x,1)=";" then
					k=x
					codart=mid(dd,iniz2,k-iniz2)
					iniz2=k+1
					exit for
				end if
			next%>
			<p align="center"><input type=text name="<%=codart%>" value="<%=mid(dd1,i,1)%>" maxlength=1></td>
		</tr>
		<%next
		session("qta")=dd1
		session("artcar")=numart
		session("stringa")=dd
		'response.write dd1%>
	</table>
	</td>
	<input type=submit name="controlla" value="controlla">
	<input type=submit name="svuota" value="svuota">
	<input type=hidden name="svuotaele" value="">
	</form>
	<%else
		response.write "il carrello è vuoto"
	end if
else
	response.write "è necessario registrarsi per poter usufruire del carrello"
end if
%>
</body>
</html>