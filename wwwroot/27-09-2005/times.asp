<%@ Language=JavaScript %>
<%
	
	 if (Request.QueryString("Province").count == 0 & Request.QueryString("TEAM").count == 0)
	 {%>
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="en-za">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE>Times - Province</TITLE>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="639" id="table1">
<tr><td><p align="center"><b>
<font size="6" color="#000099" face="Arial Rounded MT Bold">Times</font></b></p><div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="304" id="table2" style="font-weight: bold">
<tr><td colspan="2" height="31"><p align="center"><b>Select your province</b></td>
</tr><tr height="25"><td><font color="#000099"><b>Province Name</b></font></td>
<td><font color="#000099"><b>Code</b></font></td></tr>
<tr><td><a href="times.asp?Province=WP">Western Province</a></td>
<td>WP</td></tr>
</table></div></td></tr></table></BODY></HTML><% }else
	{
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	if(Request.QueryString("TEAM").count == 0)
	{
	var LSC = Request.QueryString("Province");
	rs = oConn.Execute("SELECT TEAM.Team, TEAM.TCode, TEAM.TName FROM TEAM WHERE (((TEAM.LSC)='" + LSC + "')) ORDER BY TEAM.TName;");
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="en-za">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE>Times - <%=LSC%></TITLE>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="640" id="table3" height="61">
<tr><td><div align="center"><p><b>
<font size="6" color="#000099" face="Arial Rounded MT Bold">Times - <%=LSC%></font></b></p>
<table border="0" cellpadding="0" cellspacing="0" width="304" id="table4"><tr>
<td colspan="2" height="31"><p align="center"><b>Select your Club</b></td></tr>
<tr height="25"><td><font color="#000099"><b>Club Name</b></font></td>
<td><font color="#000099"><b>Code</b></font></td></tr><%
while(!rs.eof)
{%>
<tr><td><a href="times.asp?TEAM=<%=Server.HTMLEncode(rs.Fields("Team").Value)%>&TName=<%=Server.HTMLEncode(rs.Fields("TName").Value)%>"><b><%=Server.HTMLEncode(rs.Fields("TName").Value)%></b></a></td>
<td><b><%=Server.HTMLEncode(rs.Fields("TCode").Value)%></b></td></tr><%
rs.MoveNext();
}  rs.close();
%></table></div><p>&nbsp;</td></tr></table></BODY></HTML>
<%
}else{
var Team = Request.QueryString("TEAM");
var TName = Request.QueryString("TName");
rs = oConn.Execute("SELECT Athlete.Athlete, Mid([Last],1,1) AS Alpha, Athlete.Last, Athlete.First, Athlete.Sex, Int(DateDiff('d',[Athlete]![Birth],Date())/365.25) AS Age FROM Athlete WHERE (((Athlete.Team1)=" + Team +") AND Athlete.Group = 'A')  ORDER BY Athlete.Last, Athlete.First;");
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="en-za">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE>Times - <%=TName%></TITLE>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="639" height="84">
<tr><td valign="top"><p align="center"><b>
<font size="6" color="#000099" face="Arial Rounded MT Bold"><%=TName%></font></b></p>
<div align="center">
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="400" id="table5">
<tr align="center">
<td  width ="15"><a href="#A">A</a></td><td  width ="15"><a href="#B">B</a></td><td width="15"><a href="#C">C</a></td><td width="15"><a href="#D">D</a></td>
<td width="15"><a href="#E">E</a></td><td width="15"><a href="#F">F</a></td><td width="15"><a href="#G">G</a></td><td width="15"><a href="#H">H</a></td>
<td width="15"><a href="#I">I</a></td><td width="15"><a href="#J">J</a></td><td width="15"><a href="#K">K</a></td><td width="15"><a href="#L">L</a></td>
<td width="15"><a href="#M">M</a></td><td width="15"><a href="#N">N</a></td><td width="15"><a href="#O">O</a></td><td width="15"><a href="#P">P</a></td>
<td width="15"><a href="#Q">Q</a></td><td width="15"><a href="#R">R</a></td><td width="15"><a href="#S">S</a></td><td width="15"><a href="#T">T</a></td>
<td width="15"><a href="#U">U</a></td><td width="15"><a href="#V">V</a></td><td width="15"><a href="#W">W</a></td><td width="15"><a href="#X">X</a></td>
<td width="15"><a href="#Y">Y</a></td><td width="15"><a href="#Z">Z</a></td></tr></table></div>
<table border="0" cellpadding="0" cellspacing="0" width="416" style="font-weight: bold">
<tr  height="25"><td width="150"><font color="#000099">Surname</font></td>
<td width="150"><font color="#000099">Name</font></td>
<td width="60"><font color="#000099">Gender</font></td>
<td><font color="#000099">Age</font></td></tr><%
var al = "";
while(!rs.eof)
{
 %>
<tr><td><a <% if(al != rs.Fields("Alpha").value){al = rs.Fields("Alpha").value%>name="<%=rs.Fields("Alpha")%>" <%}%>href="athlete_times.asp?Athlete=<%=rs.Fields("Athlete").Value%>&Toptimes=1"><%=Server.HTMLEncode(rs.Fields("Last").Value)%></a></td><td><%=Server.HTMLEncode(rs.Fields("First").Value)%></td><td><%if(rs.Fields("Sex").Value == "M"){%>Male<%}else{%>Female<%}%></td><td><%=Server.HTMLEncode(rs.Fields("Age").Value)%></td></tr><%
rs.MoveNext();
} rs.close(); 
%></table><p>&nbsp;</div></td></tr></table></body></HTML><%}oConn.close();}



%>