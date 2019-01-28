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
else{t = "";}
return t;}

function Age(Lo,Hi)
{
if(Lo <  Hi)
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
function fDate(d)
{
return d.getYear() + "/" + (d.getMonth() +1) + "/" + d.getDate();
}
function incYDate(d,i)
{
d.setYear((d.getYear()+1900+i));
return fDate(d);
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
	 if((Request.QueryString("Birth").count == 0 & Request.QueryString("Birth") == "") || (Request.QueryString("Age").count == 0 & Request.QueryString("Age") == "") )
	 {
	 Response.Redirect("times.asp?Athlete=" + Request.QueryString("Athlete") );
	 }else{
	 
if(Request.QueryString("Det_Age").count == 0 || Request.QueryString("Det_Age") == "")
{
rs = oConn.Execute("SELECT Meet, MName, Start, [End], Course, Location, DateAdd('yyyy',Int((DateDiff('d',DateValue('" + Request.QueryString("Birth") + "'),[MEET]![Start])/365.242199)+0.002),DateValue('" + Request.QueryString("Birth") + "')) As Def_Age,Int((DateDiff('d',DateValue('" + Request.QueryString("Birth") + "'),[MEET]![Start])/365.242199)+0.002) as Age FROM MEET WHERE ((([Start]) > Date())) ORDER BY Start ASC;");
var dat = new Date();
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="en-za">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE>Athlete Time Standards Age</TITLE>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body topmargin="0" leftmargin="0">


<table border="0" cellpadding="0" cellspacing="0" style="font-weight: bold" width="800">
<tr style="font-weight: bold">
	<td colspan="5" align="center" height="47" valign="top">
	<font face="Arial Rounded MT Bold" size="6" color="#000080">Athlete Time Standards</font></td></tr>
<tr valign="top" height="25"  align="left" style="font-weight: bold; color:#000080" class="THeader">
<td colspan="5" align="center" >Please select criteria to base Athlete's age on.</td>
</tr>
<tr valign="top" height="25"  align="left" style="font-weight: bold; color:#000080" class="THeader">
<td colspan="5" align="center" >
<table border="0" cellpadding="3" cellspacing="0" width="470">
	<tr>
<td class="THeader" >Age at:</td>
<td width="91" >
<a href="athlete_qualify.asp?Athlete=<%=Request.QueryString("Athlete")%>&StdFile=<%=Request.QueryString("StdFile")%>&Det_Age=<%=fDate( new Date())%>&Course=L">
Current age</a></td>
<td width="55" >
<a href="athlete_qualify.asp?Athlete=<%=Request.QueryString("Athlete")%>&StdFile=<%=Request.QueryString("StdFile")%>&Det_Age=<%=incYDate(new Date(Request.QueryString("Birth")),(parseInt(Request.QueryString("Age")) + 1))%>&Course=L">
<%=(parseInt(Request.QueryString("Age")) + 1)%> yrs</a></td>
<td width="55" >
<a href="athlete_qualify.asp?Athlete=<%=Request.QueryString("Athlete")%>&StdFile=<%=Request.QueryString("StdFile")%>&Det_Age=<%=incYDate(new Date(Request.QueryString("Birth")),(parseInt(Request.QueryString("Age")) + 2))%>&Course=L">
<%=(parseInt(Request.QueryString("Age")) + 2)%> yrs</a></td>
<td width="55" >
<a href="athlete_qualify.asp?Athlete=<%=Request.QueryString("Athlete")%>&StdFile=<%=Request.QueryString("StdFile")%>&Det_Age=<%=incYDate(new Date(Request.QueryString("Birth")),(parseInt(Request.QueryString("Age")) + 3))%>&Course=L">
<%=(parseInt(Request.QueryString("Age")) + 3)%> yrs</a></td>
<td width="55" >
<a href="athlete_qualify.asp?Athlete=<%=Request.QueryString("Athlete")%>&StdFile=<%=Request.QueryString("StdFile")%>&Det_Age=<%=incYDate(new Date(Request.QueryString("Birth")),(parseInt(Request.QueryString("Age")) + 4))%>&Course=L">
<%=(parseInt(Request.QueryString("Age")) + 4)%> yrs</a></td>
<td width="55" >
<a href="athlete_qualify.asp?Athlete=<%=Request.QueryString("Athlete")%>&StdFile=<%=Request.QueryString("StdFile")%>&Det_Age=<%=incYDate(new Date(Request.QueryString("Birth")),(parseInt(Request.QueryString("Age")) + 5))%>&Course=L">
<%=(parseInt(Request.QueryString("Age")) + 5)%> yrs</a></td>
	</tr>
</table>
</td>
</tr>
<tr valign="top" height="25"  align="left" style="font-weight: bold; color:#000080">
<td width="300" height="17" >Meet</td>
<td width="90" height="17">Start Date</td>
<td width="90" height="17">End Date</td>
<td width="60" height="17">Course</td>
<td height="17">Location</td></tr>

<%
var Age = 0;
while(!rs.eof){if(Age != rs.Fields("Age")){Age = rs.Fields("Age").Value;%><tr>
	<td align="left"  height="29" style="border-bottom: 2px solid #6600CC; padding-bottom:2px" valign="bottom"><a href="athlete_qualify.asp?Athlete=<%=Request.QueryString("Athlete")%>&StdFile=<%=Request.QueryString("StdFile")%>&Det_Age=<%=rs.Fields("Def_Age")%>&Course=L">Age for gala's <%=rs.Fields("Age")%></a></td>
	<td colspan="4"></td></tr><%}%><tr><td><%=rs.Fields("MName")%></td><td><%=rs.Fields("Start")%></td><td><%=rs.Fields("End")%></td><td align="center"><%=rs.Fields("Course")%></td><td><%=rs.Fields("Location")%></td></tr>
<%rs.MoveNext()}%>
</table>
<p>&nbsp;</p>
</body></HTML>




<%	 
}else{
	if(Request.QueryString("StdFile").count == 0 || Request.QueryString("StdFile") == "" || Request.QueryString("Course").count == 0 || Request.QueryString("Course") == "")
	{
	rs = oConn.Execute("SELECT StdFile, Descript, Year FROM STDNAME ORDER BY StdFile;");
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Athlete Time Standards Selection</title>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body  topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="657" style="font-weight: bold">
<tr><td valign="top"><div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="432" style="font-weight: bold">
<tr><td colspan="3" align="center" height="53" valign="bottom">
	<font face="Arial Rounded MT Bold" size="6" color="#000080">Athlete Time Standards</font></td></tr>
<tr><td colspan="3" align="center" height="26" valign="bottom">
	Please select a Qualifying Standard</td></tr>
<tr><td width="114" height="11"></td><td width="188" height="11"></td>
	<td width="58" height="11"></td>
</tr>
<%while(!rs.eof){%>
<tr><td><a href="athlete_qualify.asp?Athlete=<%=Request.QueryString("Athlete")%>&StdFile=<%=rs.Fields("StdFile")%>&Det_Age=<%=Request.QueryString("Det_Age")%>&Course=L"><%=rs.Fields("StdFile")%></a></td><td><%=rs.Fields("Descript")%></td><td><%=rs.Fields("Year")%></td></tr><%rs.MoveNext()}%></table><p>&nbsp;</td></tr></table>
</table>
<p><br>&nbsp;</p>
</body></html><%rs.close();
oConn.close();





}
else{
var StdFile = Request.QueryString("StdFile");
var Athlete = Request.QueryString("Athlete");
var Course = Request.QueryString("Course");
	rs = oConn.Execute("SELECT Athlete.Athlete, Athlete.Last, Athlete.First, Athlete.Initial, Athlete.Sex, Athlete.Birth, Int((DateDiff('d',[Athlete]![Birth],Date())/365.242199)+0.002) AS Age ,  Int((DateDiff('d',[Athlete]![Birth],'" + Request.QueryString("Det_Age") + "')/365.242199)+0.002) AS Det_Age FROM Athlete WHERE (((Athlete.Athlete)=" + Athlete + ")) ORDER BY Athlete.Last, Athlete.First;");%>
<html>
<HEAD>
<meta http-equiv="Content-Language" content="en-za">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE>Athlete T/S - <%=StdFile%> of <%=rs.Fields("Last").Value%>, <%=rs.Fields("First").Value%> (<%=rs.Fields("Det_Age")%>yrs) - <%=Course%>C</TITLE>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>

&nbsp;<STYLE TYPE="text/css">
.goal{color: #6600CC;padding-left:15px; font-size:10pt}
.dashborder {border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: dashed; border-bottom-width: 1px;padding-top:0px;padding-bottom:4px}</STYLE></head><body topmargin="0" leftmargin="0"><table border="0" cellpadding="0" cellspacing="0" width="640" id="table1" style="font-weight: bold" height="102">
<tr><td valign="top"><div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="531" id="table2" style="font-weight: bold">
<tr><td><%=Server.HTMLEncode(rs.Fields("Last").Value)%>, <%=Server.HTMLEncode(rs.Fields("First").Value)%></td>
<td width="174" align="right"><a href="athlete_times.asp?Athlete=<%=Athlete%>">All Times</a></td></tr>
<tr><td><%if(rs.Fields("Sex").Value == "M"){%>Male<%}else{%>Female<%}%></td>
<td width="174" align="right"><a href="athlete_times.asp?Athlete=<%=Athlete%>&Toptimes=1">Top Times only</a></td></tr>
<tr><td><%=Server.HTMLEncode(rs.Fields("Birth").Value)%>&nbsp;&nbsp; (<%=Server.HTMLEncode(rs.Fields("Age").Value)%>yrs 
	- Currently)&nbsp; </td>
<td width="174" align="right"><a href="athlete_qualify.asp?Athlete=<%=Athlete%>&Age=<%=Server.HTMLEncode(rs.Fields("Age").Value)%>&Birth=<%=Server.HTMLEncode(rs.Fields("Birth").Value)%>">Qualifying Time Standards</a></td></tr>
<tr><td>Determining date,&nbsp; <%=Request.QueryString("Det_Age")%>&nbsp;&nbsp; (<%=Server.HTMLEncode(rs.Fields("Det_Age").Value)%>yrs)</td>
<td width="174" align="right"><font color="#FFFFFF">Change</font>
<a href="athlete_qualify.asp?Athlete=<%=Athlete%>&Age=<%=rs.Fields("Age").Value%>&Birth=<%=rs.Fields("Birth").Value%>&StdFile=<%=StdFile%>&Course=L">Age</a>
<font color="#FFFFFF">or</font>
<a href="athlete_qualify.asp?Athlete=<%=Athlete%>&StdFile=&Det_Age=<%=Request.QueryString("Det_Age")%>">Standard</a></td></tr>
<tr><td align="right">
</td>
<td width="174" align="right"><%if(Course=="L"){%>
<a href="athlete_qualify.asp?Athlete=<%=Athlete%>&StdFile=<%=StdFile%>&Det_Age=<%=Request.QueryString("Det_Age")%>&Course=S">Course &nbsp;SC</a><%}else{%>
<a href="athlete_qualify.asp?Athlete=<%=Athlete%>&StdFile=<%=StdFile%>&Det_Age=<%=Request.QueryString("Det_Age")%>&Course=L">Course &nbsp;LC</a><%}%></td></tr></table>

<%
rs.Close();
rs = oConn.Execute("SELECT * FROM STDNAME WHERE StdFile='" + StdFile + "';");

var SQL = "",Stroke ="",Distance="",begin="", bc="",end="",times="",names="",levels="",lcount = 1;
if(Course == "L")
{
colom = 12;
}else
{
colom = 0;
}
if(rs.Fields(3) != "")
{
names =",'" + rs.Fields(3).Value + "'";
times = ",[S(" + colom + ")]";
levels ="IIf(IsNull([S(" + (colom) + ")]),True,[t].[SCORE]<[S(" + (colom) + ")]) And IIf(IsNull([F(" + (colom) + ")]),True,[t].[SCORE]>[F(" + (colom) + ")])," + lcount + ",";
lcount++;
pos = 4;
while(rs.Fields(pos) != "")
{
names +=",'" + rs.Fields(pos).Value + "'";
times += ",[S(" + (colom + pos -3) + ")]";
levels +="IIf(IsNull([S(" + (colom + pos -3) + ")]),True,[t].[SCORE]<[S(" + (colom + pos -3) + ")]) And IIf(IsNull([F(" + (colom + pos -3) + ")]),True,[t].[SCORE]>[F(" + (colom + pos -3) + ")])," + lcount + ",";
lcount++;
pos++;
}

SQL = "SELECT t.Lo_Age, t.Hi_age, t.Distance, t.Stroke, Choose([Level]" + names + ",'- - -')  AS STD, Switch(" + levels + "[t].[SCORE]," + lcount + ") AS [Level] , Choose([Level]" + times + ") AS [Current Level] , Choose([Level]-1" + names + ") AS [NNext Level], Choose(([Level]-1)" + times + ") AS [Next Level], [t].[SCORE] - Choose(([Level]-1)" + times + ") AS Improve, t.F_P, t.SCORE,MEET.[End], MEET.MName"
+" FROM (SELECT st.*, rs.MEET, rs.SCORE, rs.F_P"
+" FROM (SELECT RESULT.MEET,RESULT.F_P, RESULT.STROKE, RESULT.DISTANCE,RESULT.SCORE FROM RESULT WHERE (((RESULT.RANK)=1) AND ((RESULT.ATHLETE)="+Athlete+") AND ((RESULT.COURSE)='" + Course + "'))) AS rs, (SELECT [" + StdFile + "].* FROM [" + StdFile + "], Athlete"
+" WHERE ((([" + StdFile + "]![Lo_Age])<=Int((DateDiff('d',[Athlete]![Birth],DateValue('" + Request.QueryString("Det_Age") + "'))/365.242199)+0.002)) AND (([" + StdFile + "]![Hi_age])>=Int((DateDiff('d',[Athlete]![Birth],DateValue('" + Request.QueryString("Det_Age") + "'))/365.242199)+0.002)) AND (([" + StdFile + "]![I_R])='I') AND (([" + StdFile + "]![Sex])=[Athlete]![Sex]) AND ((Athlete.Athlete)="+Athlete+"))) AS st"
+" WHERE ((([rs].[STROKE] & [rs].[DISTANCE])=[st].[Stroke] & [st].[Distance]))) AS t INNER JOIN MEET ON t.MEET = MEET.Meet"
+"  ORDER BY t.Stroke, t.Distance;";



//Response.write(":"+SQL);
rs.close();
rs = oConn.Execute(SQL);
}
%>
<table border="0" cellpadding="0" cellspacing="0" width="600" id="table3" style="font-weight: bold">
<%
var goal;
if(!rs.eof)
{ 
if(rs.Fields("Current Level").Value == null & rs.Fields("Level").Value == 1)
{%><tr><td colspan="6" align="center">Sorry NO Time Standards found in this Course</td></tr><%}else{%>
<tr><td valign="bottom" colspan="6" height="30" align="center" ><font  size="4" color="#000080"><u>
 <%=StdFile%>&nbsp;<%=Age(rs.Fields("Lo_Age").Value ,rs.Fields("Hi_age").Value )%>&nbsp;<%if(Course=="L"){%>
Long Course 50m<%}else{%>
Short Course 25m</a><%}%></u></font></td></tr>
<tr height="10"><td width="75"></td><td width="64">q/stand</td><td width="81">
	q/time</td><td width="55"></td><td width="91"></td><td width="234"></td></tr>
<%
while(!rs.eof)
{
if(rs.Fields("STROKE").Value != Stroke || rs.Fields("DISTANCE").Value != Distance)
{
Stroke = rs.Fields("STROKE").Value;
Distance = rs.Fields("DISTANCE").Value;
if(rs.Fields("NNext Level").Value !=null)
	{
	goal = "<div class='goal'>My Goal Time is,&nbsp;&nbsp;" + rs.Fields("NNext Level").Value + "&nbsp;-&nbsp;" + CTime(rs.Fields("Next Level").Value) + "&nbsp; must improve by&nbsp;" + CTime(rs.Fields("Improve").Value) + "s</div>";
//goal = "<tr height ='20px' valign='top' align='left' style='padding-top: 6px'><td colspan='10'>My Goal Time is,&nbsp;&nbsp;" + rs.Fields("NNext Level").Value + "&nbsp;-&nbsp;" + CTime(rs.Fields("Next Level").Value) + "&nbsp; must improve by&nbsp;" + CTime(rs.Fields("Improve").Value) + "s</td></tr>";
}else{goal="&nbsp;";}
 %><tr  height="30" valign="bottom">
	<td colspan="3" class='dashborder' height="33"><%=Distance%>&nbsp;<%=strokes[Stroke] %></td>
	<td class='dashborder' colspan="3" height="33" align='left'><%=goal%></td></tr><%}
	
%>
<tr><td><%=CTime(rs.Fields("SCORE").Value)%></td><td><%=rs.Fields("STD").Value%></td><td>&nbsp;<%if(rs.Fields("Current Level").Value != null ){Response.write(CTime(rs.Fields("Current Level").Value))}%></td><td><%if(rs.Fields("F_P").Value=="F"){%>Final<%}else{%>Prelim<%}%></td><td><%=Server.HTMLEncode(rs.Fields("End"))%></td>
	<td width="234"><%=Server.HTMLEncode(rs.Fields("MName"))%></td></tr><%
rs.MoveNext();}
rs.close();
}%>
<%
}else{%><tr><td colspan="6" align="center">Sorry Athlete doesn't qualify for anything.</td></tr>
<%}
oConn.close();  
%></table><p><br>&nbsp;</p>
</body></html>
<%
}}}}
%>