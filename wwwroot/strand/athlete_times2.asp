<%@ Language=JavaScript %>
<%
Response.CacheControl = "Public";
Response.Expires = 1440;
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
else{t = "";}//dq
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
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	 if (Request.QueryString("Athlete").count == 0)
	 {
	 
	 Response.Redirect("times.asp");
	 }
	 else
	 {
	var Athlete = Request.QueryString("Athlete");
	rs = oConn.Execute("SELECT Athlete.Athlete, Athlete.Last, Athlete.First, Athlete.Initial, Athlete.Sex, Athlete.Birth, Int((DateDiff('d',[Athlete]![Birth],Date())/365.242199)+0.002) AS Age FROM Athlete WHERE (((Athlete.Athlete)=" + Athlete + ")) ORDER BY Athlete.Last, Athlete.First;");%>
<html>
<HEAD>
<meta http-equiv="Content-Language" content="en-za">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE><%if(Request.QueryString("Toptimes").count == 1){%>Top Times<%}else{%>All Times<%}%> - <%=rs.Fields("Last").Value%>, <%=rs.Fields("First").Value%></TITLE>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


<STYLE TYPE="text/css">
.dashborder {border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: dashed; border-bottom-width: 1px}</STYLE>
</head>
<body topmargin="0" leftmargin="0">


<table border="0" cellpadding="0" cellspacing="0" width="640" id="table1" style="font-weight: bold" height="102">
<tr><td valign="top"><div align="center">


<table border="0" cellpadding="0" cellspacing="0" width="531" id="table2" style="font-weight: bold">
<tr><td><%=Server.HTMLEncode(rs.Fields("Last").Value)%>, <%=Server.HTMLEncode(rs.Fields("First").Value)%></td>
<td width="175" align="right"><%if(Request.QueryString("Toptimes").count == 1){%><a href="athlete_times.asp?Athlete=<%=Athlete%>">All Times</a><%}else{%><a href="athlete_times.asp?Athlete=<%=Athlete%>&Toptimes=1">Top Times only</a><%}%></td></tr>
<tr><td><%if(rs.Fields("Sex").Value == "M"){%>Male<%}else{%>Female<%}%></td>
<td width="175" align="right"><a href="athlete_qualify.asp?Athlete=<%=Athlete%>&Age=<%=Server.HTMLEncode(rs.Fields("Age").Value)%>&Birth=<%=Server.HTMLEncode(rs.Fields("Birth").Value)%>&StdFile=&Course=L">Qualifying Time Standards</a></td></tr>
<tr><td height="25"><%=Server.HTMLEncode(rs.Fields("Birth").Value)%>&nbsp;(<%=Server.HTMLEncode(rs.Fields("Age").Value)%>yrs)</td>
<td height="25" width="175">&nbsp;</td></tr></table><%
rs.Close();
if(Request.QueryString("Toptimes").count == 0)
{
var Stroke, Distance, Course;
rs = oConn.Execute("SELECT RESULT.MEET, RESULT.COURSE, RESULT.MTEVENT, RESULT.STROKE, RESULT.DISTANCE, RESULT.SCORE, RESULT.F_P, MEET.[End], MEET.MName FROM (RESULT INNER JOIN Athlete ON RESULT.ATHLETE = Athlete.Athlete) INNER JOIN MEET ON RESULT.MEET = MEET.Meet WHERE (((RESULT.ATHLETE)=" + Athlete + ")) ORDER BY RESULT.COURSE, RESULT.STROKE, RESULT.DISTANCE, RESULT.SCORE;");%>
<table border="0" cellpadding="0" cellspacing="0" width="531" id="table3" style="font-weight: bold">
<tr height="0"><td width="75"></td><td width="76"></td><td width="104"></td><td width="276"></td></tr><%
if(!rs.eof)
{
while(!rs.eof)
{
if(rs.Fields("COURSE").Value != Course)
{
Course = rs.Fields("COURSE").Value;
%><tr><td colspan="4" height="30" valign="bottom" align="center">
<font  size="4" color="#000080"><u><%if(rs.Fields("COURSE") == "S"){%>Short Course - 25m<%}else{%>Long Course - 50m<%}%></u></font></td></tr><%
}
if(rs.Fields("STROKE").Value != Stroke || rs.Fields("DISTANCE").Value != Distance)
{
Stroke = rs.Fields("STROKE").Value;
Distance = rs.Fields("DISTANCE").Value;
 %><tr><td valign="bottom" colspan="4" height="30" class=dashborder ><%=Distance%>&nbsp;<%=strokes[Stroke] %></td></tr><tr><%}
%>
<tr><td><%=CTime(rs.Fields("SCORE").Value)%></td><td><%if(rs.Fields("F_P").Value=="F"){%>Final<%}else{%>Prelim<%}%></td><td><%=Server.HTMLEncode(rs.Fields("End"))%></td><td><a href='meets/meet.asp?Meet=<%=rs.Fields("MEET")%>&MtEvent=<%=rs.Fields("MTEVENT")%>&MName=<%=Server.HTMLEncode(rs.Fields("MName"))%>'><%=Server.HTMLEncode(rs.Fields("MName"))%></a></td></tr><%
rs.MoveNext();
}}else{%><tr><td colspan="4" align="center">Sorry no Times Found for Athlete</td></tr>
<%}  
%></table><%}else{

var Course,stroke,distance;
rs = oConn.Execute("SELECT RESULT.MEET, RESULT.COURSE, RESULT.MTEVENT, RESULT.STROKE, RESULT.DISTANCE, RESULT.SCORE, RESULT.F_P, MEET.[End], MEET.MName FROM (RESULT INNER JOIN Athlete ON RESULT.ATHLETE = Athlete.Athlete) INNER JOIN MEET ON RESULT.MEET = MEET.Meet WHERE ((RESULT.RANK = 1) AND ((RESULT.ATHLETE)=" + Athlete + ")) ORDER BY RESULT.COURSE, RESULT.STROKE, RESULT.DISTANCE, RESULT.SCORE, MEET.[End];");%>
<table border="0" cellpadding="0" cellspacing="0" width="608" id="table3" style="font-weight: bold">
<tr height="0"><td width="70"></td><td width="40"></td><td width="75"></td><td width="60"></td><td width="90"></td><td></td></tr><%
if(!rs.eof)
{
while(!rs.eof)
{
if(rs.Fields("COURSE").Value != Course)
{
Course = rs.Fields("COURSE").Value;
%><tr><td colspan="6" height="40" valign="middle" align="center"><font  size="4" color="#000080"><u><%if(rs.Fields("COURSE") == "S"){%>Short Course - 25m<%}else{%>Long Course - 50m<%}%></u></font></td></tr><%
}%>
<tr><td><%=CTime(rs.Fields("SCORE").Value)%></td><td><%=rs.Fields("DISTANCE").Value%></td><td><%=strokes[rs.Fields("STROKE").Value]%></td><td><%if(rs.Fields("F_P").Value=="F"){%>Final<%}else{%>Prelim<%}%></td><td><%=Server.HTMLEncode(rs.Fields("End"))%></td><td><a href='meets/meet.asp?Meet=<%=rs.Fields("MEET")%>&MtEvent=<%=rs.Fields("MTEVENT")%>&MName=<%=Server.HTMLEncode(rs.Fields("MName"))%>'><%=Server.HTMLEncode(rs.Fields("MName"))%></a></td></tr><%
rs.MoveNext();
}}else{%><tr><td colspan="7" align="center">Sorry no Times Found for Athlete</td></tr>
<%}  %></table><%}%></td></tr></table><p>&nbsp;</p></body></html>
<%
}
rs.close();
oConn.close();
%>