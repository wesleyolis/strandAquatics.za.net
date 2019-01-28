<%@ Language=JavaScript %>
<!--METADATA TYPE="typelib" 
uuid="00000206-0000-0010-8000-00AA006D2EA4" -->
<%

Response.CacheControl = "Public";
Response.Expires = -1;

function CTime (Score)
{
if(Score !=0)
{
if(Score<6000)
{t = (Score/100);
if(t <10.0)
{t ="0"+t;}
if(t.toString().length == 2)
{t = t + ".00";}else{
if(t.toString().length<5)
{t=t+"0";}}
}
else
{
s = (Score % 100);
if(s==0){s=".00"}else{
if(s <10){s = "0" + s;}}
if(s.toString().length<2){s = s + "0";}
Score = ((Score -s) / 100);
m = (Score % 60);
if(m < 10)
{m = "0" + m;}
Score = ((Score - m)/60);
h = (Score % 60);
if(h.toString().length<2){h = "0" + h;}
t = h + ":" + m + "." + s;}}
else{t = "D/Q";}
return t;}

function Age(Lo,Hi)
{if(Lo <  Hi)
{
if(Lo > 0 & Hi < 99)
{
age = Lo + " - " + Hi + "yrs";
}else
{
if(Hi == 99 & Lo >0)
{
age = Lo + "&nbsp;&amp;&nbsp;Over";
}
else
{
if(Lo == 0 & Hi < 99)
{
age = Hi + "&nbsp;&amp;&nbsp;Under";
}
else
{
if(Lo == 0 & Hi == 99)
{
age = "Open";
}}}}}else{age = Hi + "yrs";}

return age;
}

var strokes = new Array("","Freestyle","Back","Breast","Butterfly","Medley");

var Course,Lo,Hi,Stroke,Distance,Gender,Province,Page;
Province = Request.QueryString("Province");
Course = Request.QueryString("Course");
Lo = Request.QueryString("Lo");
Hi = Request.QueryString("Hi");
Stroke = Request.QueryString("Stroke");
Distance = Request.QueryString("Distance");
Gender = Request.QueryString("Gender");
Page = Request.QueryString("Page");
if(Hi < Lo || Lo > 98 || Hi > 99 || Lo < 0 || Hi < 0)
{Response.Redirect("rankings.asp");}
%>
<%if(Request.QueryString("Province").count==0 || Province == ""){%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rankings</title>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body  topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="793">
<tr><td align="center">
<div class="heading">Rankings</div><br>
&nbsp;<div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="304">
<tr><td colspan="2" height="31" align="center" class="Instruct">Select your province</td>
</tr><tr class="THeader"><td width="246"><font color="#000099"><b>Province Name</b></font></td>
<td width="58"><font color="#000099"><b>Code</b></font></td></tr>
<tr><td width="246"><a href="rankings.asp?Province=All">All Province&nbsp; (South Africa)</a></td>
	<td width="58">All</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=WP">Western Province</a></td>
	<td width="58">WP</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=BO">Border</td>
	<td width="58">BO</a></td></tr>
<tr><td width="246"><a href="rankings.asp?Province=CG">Central Guateng</a></td>
	<td width="58">CG</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=EP">Eastern Province</a></td>
	<td width="58">EP</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=ES">Eastern Swimming</a></td>
	<td width="58">ES</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=FS">Freestate</a></td>
	<td width="58">FS</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=GW">Griqualand West</a></td>
	<td width="58">GW</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=KZ">Kwa-Zulu Natal</a></td>
	<td width="58">KZ</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=MP">Mpumalanga</a></td>
	<td width="58">MP</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=NF">Northern Freestate</a></td>
	<td width="58">NF</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=NN">Northern Natal</a></td>
	<td height="18" width="58">NN</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=NP">Northern Province</a></td>
	<td width="58">NP</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=NT">Northern Tigers</a></td>
	<td width="58">NT</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=NW">North West</a></td>
	<td width="58">NW</td></tr>
<tr><td width="246"><a href="rankings.asp?Province=VT">Vaal Triangle</a></td>
	<td width="58">VT</td></tr>
</table></div>
<p><br></p>

<%}else{if(Request.QueryString("Course").count==0 || Course == ""){%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=Province%> Rankings</title>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body  topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="600" style="font-weight: bold">
<tr><td align="center">
<div class="heading"><%=Province%> Rankings</div><br><br>
<div class="Instruct">Please select a Course</div>
<p><a href="rankings.asp?Province=<%=Province%>&Course=L">Long Course 50m</a><a href="rankings.asp?Province=<%=Province%>&Course=S"><br>Short Course 25m</a><%}else
{if(Request.QueryString("Lo").count == 0 || Request.QueryString("Hi").count == 0 || Lo == "" || Hi == "")
{%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=Province%> Rankings - <%if(Course == "L"){%>Long Course 50m<%}else{%>Short Course 25m<%}%></title>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body  topmargin="0" leftmargin="0">
</p>
<table border="0" cellpadding="0" cellspacing="0" width="793" style="font-weight: bold">
<tr><td align="center">
<div class="heading"><%=Province%> Rankings</div><br>
<font color="#000080"><%if(Course == "L"){%>Long Course 50m<%}else{%>Short Course 25m<%}%></font><br><br>
<div class="Instruct">
Please select&nbsp; an Age Group</div><br><br><table border="0" cellpadding="0" cellspacing="0" width="300">
<tr><td width="176"><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=0&Hi=99&Page=1">Open&nbsp; (all ages)</a></td>
<td width="124"><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=19&Hi=19&Page=1">19 yrs</a></td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=21&Hi=99&Page=1">21 &amp; Over</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=18&Hi=18&Page=1">18 yrs</a></td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=18&Hi=21&Page=1">18 - 21</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=17&Hi=17&Page=1">17 yrs</a></td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=19&Hi=99&Page=1">19 &amp; over</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=16&Hi=16&Page=1">16 yrs </a> </td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=17&Hi=18&Page=1">17 - 18</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=15&Hi=15&Page=1">15 yrs </a> </td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=17&Hi=99&Page=1">17 &amp; Over</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=14&Hi=14&Page=1">14 yrs </a> </td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=16&Hi=17&Page=1">16-17</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=13&Hi=13&Page=1">13 yrs </a> </td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=15&Hi=18&Page=1">15 - 18</a></td>
<td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=12&Hi=12&Page=1">12 yrs</font></a></td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=15&Hi=16&Page=1">15-16</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=0&Hi=12&Page=1">12 &amp; under</a></td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=15&Hi=99&Page=1">15 &amp; Over</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=11&Hi=11&Page=1">11 yrs</a></td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=13&Hi=14&Page=1">13 - 14</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=10&Hi=10&Page=1">10 yrs</a></td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=13&Hi=16&Page=1">13 - 16</a></td><td>
<a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=0&Hi=10&Page=1">10 &amp; under</a></td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=11&Hi=12&Page=1">11- 12</a></td>
<td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=9&Hi=9&Page=1">9 yrs </a> </td></tr>
<tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=11&Hi=14&Page=1">11 -14</a></td>
<td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=0&Hi=8&Page=1">8 &amp; under</a> </td></tr>
</table><%}else{
var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);

if(Request.QueryString("Stroke").count==0 || Stroke == "" || Request.QueryString("Distance").count==0 || Distance == "" || Request.QueryString("Gender").count==0 || Gender == "")
{
	rs = oConn.Execute("SELECT Sex, STROKE, DISTANCE FROM ResultInfo WHERE (((COURSE)='" + Course + "')  AND [Birth] >=" + Lo + " And [Birth] <=" + Hi + ") GROUP BY Sex, STROKE, DISTANCE ORDER BY Sex, STROKE, DISTANCE;");
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=Province%> Rankings -  <%if(Course == "L"){%>LC<%}else{%>SC<%}%>, <%=Age(Lo,Hi)%></title>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body  topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="600" style="font-weight: bold">
<tr><td align="center">
<div class="heading"><%=Province%> Rankings</div><br>
<font color="#000080"><%if(Course == "L"){%>Long Course 50m<%}else{%>Short Course 25m<%}%><br><%=Age(Lo,Hi)%></font><br><br>
<table border="0" cellpadding="0" cellspacing="0" width="470">
<tr align="center" class="THeader"><td width="50%">Female</td>
<td width="50%">Male</td></tr>
<td align="center" valign="top"><%if(rs.eof){%><tr><td colspan="6"  align="center">
<font size="3" color="#000000"><b>No results found in this Category! :'(</b></font></td></tr><%}else{%>
<table border="0" cellpadding="0" cellspacing="0" style="font-weight: bold">
<%while(!rs.eof & rs.Fields("Sex")=="F"){%><tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=<%=Lo%>&Hi=<%=Hi%>&Stroke=<%=rs.Fields("STROKE")%>&Distance=<%=rs.Fields("DISTANCE")%>&Gender=F"><%=rs.Fields("DISTANCE")%>m&nbsp&nbsp<%=strokes[rs.Fields("STROKE")]%></a></td></tr>
<%rs.MoveNext()}%></table></td><td align="center" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="font-weight: bold">
<%while(!rs.eof){%><tr><td><a href="rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=<%=Lo%>&Hi=<%=Hi%>&Stroke=<%=rs.Fields("STROKE")%>&Distance=<%=rs.Fields("DISTANCE")%>&Gender=M&Page=1"><%=rs.Fields("DISTANCE")%>m&nbsp&nbsp<%=strokes[rs.Fields("STROKE")]%></a></td></tr>
<%rs.MoveNext()}}%>
</table></td></tr></table>
<%}else{
var SQLprov = "";
if(Province != "All")
{
SQLprov = "AND ([TEAM].[LSC] = '" + Province + "')";
}
else
{
SQLprov = "";
}

rs = Server.CreateObject("ADODB.Recordset");
rs.Open ("SELECT RESULT.SCORE, Athlete.Last, Athlete.First, Int(DateDiff('d',[Athlete]![Birth],Date())/365.25) AS Ath_Age, [TEAM].[TCode] & '-' & [TEAM].[LSC] AS Team, First(MEET.MName) AS MName FROM ((RESULT INNER JOIN Athlete ON RESULT.ATHLETE = Athlete.Athlete) INNER JOIN TEAM ON Athlete.Team1 = TEAM.Team) INNER JOIN MEET ON RESULT.MEET = MEET.Meet WHERE (((RESULT.TRANK)=1) " + SQLprov + " AND (Int(DateDiff('d',[Athlete]![Birth],Date())/365.25)>=" + Lo + " And Int(DateDiff('d',[Athlete]![Birth],Date())/365.25)<="+ Hi +") AND ((RESULT.DISTANCE)="+Distance+") AND ((RESULT.STROKE)="+Stroke+") AND ((RESULT.COURSE)='"+Course+"') AND ((RESULT.I_R)= 'I') AND ((Athlete.Sex)='"+Gender+"') AND ((Athlete.Group)='A')) GROUP BY RESULT.SCORE, Athlete.Last, Athlete.First, Int(DateDiff('d',[Athlete]![Birth],Date())/365.25), [TEAM].[TCode] & '-' & [TEAM].[LSC] ORDER BY RESULT.SCORE;", oConn, adOpenStatic);
rs.PageSize = 200;
if(rs.PageCount !=0 & (Request.QueryString("Page").count==0 || Page == "" || rs.PageCount  < Page || 10 < Page))
{
Response.Redirect("rankings.asp?Province=" + Province + "&Course=" + Course + "&Lo=" + Lo + "&Hi=" + Hi + "&Stroke=" + Stroke + "&Distance=" + Distance + "&Gender=" + Gender + "&Page=1");
}
if(rs.PageCount !=0)
{
rs.AbsolutePage = Page;
}
posit = (200 * (Page-1)) + 1;
%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=Province%> Rankings -  <%if(Course == "L"){%>LC<%}else{%>SC<%}%>, <%=Age(Lo,Hi)%> - <%=strokes[Stroke]%>, <%=Distance%>m <%if(Gender == "M"){%>Male<%}else{%>Female<%}%></title>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body  topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="600" style="font-weight: bold">
<tr><td align="center">
<div class="heading"><%=Province%> Rankings</div><br>
<font color="#000080"><%if(Course == "L"){%>Long Course 50m<%}else{%>Short Course 25m<%}%><br><%if(Gender == "M"){%>Male<%}else{%>Female<%}%>&nbsp;&nbsp;<%=Age(Lo,Hi)%>&nbsp;&nbsp;<%=Distance%>m&nbsp;&nbsp;<%=strokes[Stroke]%></font><br><br>
<%  if (rs.PageCount > 1) { %>
<table border="0" cellpadding="2" id="table2"  cellspacing="0">
<tr align="center">
<%for(p = 0;p<(rs.PageCount - 1) & p < 10;p++)
{%><td width="70"><a href='rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=<%=Lo%>&Hi=<%=Hi%>&Stroke=<%=Stroke%>&Distance=<%=Distance%>&Gender=<%=Gender%>&Page=<%=(p+1)%>'><%=((p *200)+1)%> - <%=(((p + 1) * 200)+1)%></a></td><%}%>
<td width="70"><a href='rankings.asp?Province=<%=Province%>&Course=<%=Course%>&Lo=<%=Lo%>&Hi=<%=Hi%>&Stroke=<%=Stroke%>&Distance=<%=Distance%>&Gender=<%=Gender%>&Page=<%=rs.PageCount%>'><%=((p *200)+1)%> +</a></td>
</tr>
</table>
<%}%>


<table border="0" cellpadding="0" cellspacing="0" width="789" id="table1" style="font-weight: bold; font-size:10pt">
<tr class="THeader"><td width="51">Rank</td><td width="67">Time</td>
<td width="140">&nbsp;</td><td width="140"></td>
<td width="30">&nbsp;</td><td  width="60"></td><td></td></tr><%if(rs.eof){%><tr><td colspan="7"  align="center">
<font size="3" color="#000000">No results found in this Category! :'(</font></td></tr><%}for (j = 0; j < rs.PageSize; j++){if (rs.EOF){break;}%>
<tr><td><%=posit%></td><td><%=CTime(rs.Fields("SCORE"))%></td><td><%=rs.Fields("Last")%></td><td><%=rs.Fields("First")%></td><td><%=rs.Fields("Ath_Age")%></td><td><%=rs.Fields("Team")%></td><td><%=rs.Fields("MName")%></td></tr><%rs.MoveNext();posit++;}%>
</table>
<div style="position: absolute; width: 237px; height: 20px; z-index: 1; left: 756px; top: 1px; float: right" align="right">
<p  align="right"><font size="2"><%=Date()%><br>Page Last Refreshed</font></p></div><%}   rs.Close();oConn.Close();}}}%>
<p><br></p></td></tr></table></body></html>
