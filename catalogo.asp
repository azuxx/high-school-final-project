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
<meta http-equiv="Page-Enter" content="blendTrans(Duration=1.0)">
<meta http-equiv="Site-Enter" content="blendTrans(Duration=1.0)">
<link rel="File-List" href="catalogo_file/filelist.xml">
<script language="VBscript">
sub categ_onclick(k)
	frm2.scelta.value=k
	frm2.submit
end sub
</script>
<!--[if gte mso 9]>
<xml><o:shapedefaults v:ext="edit" spidmax="1027"/>
</xml><![endif]-->
</head>

<body background="images/litbg.gif">
&nbsp;<p><!--[if gte vml 1]><v:shapetype id="_x0000_t136"
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
</v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t136" alt="Area clienti"
 style='width:237pt;height:33pt' strokeweight=".35189mm">
 <v:fill src="catalogo_file/image001.jpg" o:title="DSCN1066" color2="blue"
  recolor="t" type="frame"/>
 <v:shadow on="t" type="perspective" color="silver" opacity="52429f" origin="-.5,.5"
  matrix=",46340f,,.5,,-4768371582e-16"/>
 <v:textpath style='font-family:"Cracked Johnnie";font-size:32pt;v-text-kern:t'
  trim="t" fitpath="t" string="Catalogo prodotti"/>
</v:shape><![endif]--><![if !vml]><img border=0 width=336 height=48
src="catalogo_file/image002.gif" alt="Area clienti" v:shapes="_x0000_s1025"><![endif]><br>
<br>
<%if session("login")=false then
	response.write "Si ricorda all'utente che può consultare solo il catalogo e non può effettuare acquisti poichè non è registrato oppure non ha effettuato il log in! per registrarsi "%><a target="principale" href="registrati.htm">clicca qui</a>
<%else
	nomelog=session("nomelog")
	cognomelog=session("cognomelog")
	response.write "Cliente: " & nomelog & " " & cognomelog
end if%>
<%dim conn, rs, sql
set conn=server.createobject("ADODB.connection")
conn.open "provider=microsoft.jet.oledb.4.0; data source=D:/Andrea/HTML/negozio_online/negozio_online.mdb"
set rs=server.createobject("ADODB.recordset")
sql="Select * from Categorie"
rs.open sql, conn%>
</p>
<font face="Tahoma">Per visualizzare gli articoli di una categoria cliccare su di essa!</font>
<table border="1" width="48%" id="table1" background="images/body_bg.jpg">
	<tr>
	<th colspan=2><i><font face="Tahoma" size="5">Categorie</font></i></th>
	</tr>
<%while not rs.EOF%>
	<tr>
		<td width="16%">
		<a name="categ" id="categ" onclick='categ_onclick(<%=rs.fields("cat_id")%>)'><%=rs.fields("cat_descriz")%></a></td>
	</tr>
	<%rs.movenext
	wend%>
</table>
<%rs.close
conn.close%>
<form name="frm2" action="articoli.asp" method="post">
	<input type=hidden name="scelta" id="scelta" value="">
</form>
</body>
</html>