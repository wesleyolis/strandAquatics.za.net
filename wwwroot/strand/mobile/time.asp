<%@ Language=JavaScript %>
<% Response.ContentType = "text/vnd.wap.wml" 
%><?xml version="1.0" encoding="iso-8859-1" ?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">


<%Response.Expires = 2880%>
<%
s2=""; m2="";h2="";
function CTime (Score)
{
if(Score != "DQ" & Score != "NS")
{
s = (Score % 100);
if(s <10){s2 = "0" + s;}
else{s2 = s;}

Score = (Score -s)/100;
m = (Score % 60);
if(m < 10){m2 = "0" + m;}
else{m2 = m;}

h = (Score - m)/60;
if(h < 10){h2 = "0" + h;}
else{h2 = h;}
t = h2 + ":" + m2 + "." + s2;
}
else{t = Score}//dq
return t;}
%>



<%
//test for alpha menu chrater
if(Request.QueryString("f").Count == 0)
{
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	
	var Athlete = Request.QueryString("a");
	rs = oConn.Execute("SELECT Left(Athlete.First,1) As initial, Athlete.Last, Athlete.Sex FROM Athlete WHERE (Athlete.Athlete=" + Athlete + ")");
if(rs.eof)
{
//redirect athelete not found
%>

<wml>

<card ontimer="index.wml" title="Not Found"><timer value="30"/>
<p align="center">
Athlete not found, redirecting..
</p>
</card>
</wml>
<%
}
else
{
%>
<wml>


<card id="menu" title="<%=rs(0)%>.<%=rs(1)%>">
<p align="center"><u><%=rs(0)%>.<%=rs(1)%> - <%=rs(2)%></u></p>
<p>
<a href="time.asp?a=<%=Request.QueryString("a")%>&amp;f=l">Latest Meet Times</a><br/>
<a href="#course">Top Times</a><br/>
<a href="time.asp?a=<%=Request.QueryString("a")%>&amp;f=a">All Times</a><br/>
<a href="#qualify">Do I qualify</a><br/>
<a href="time.asp?a=<%=Request.QueryString("a")%>&amp;f=i">Athlete Info</a><br/>

<b><a href="#abbr">Abbreviations</a></b><br/>
</p>
</card>


<card id="course">
<p>
<u>Select Course</u><br/>
<select name="My_Course">
<option value="S">Short 25m Pool</option>
<option value="L">Long 50m Pool</option>
</select>
<br/>
<anchor><go method="get" href="time.asp?a=<%=Request.QueryString("a")%>&amp;f=t">
<postfield name="c" value="$(My_Course)"/></go>
Get Times</anchor>
</p>
</card>


<card id="abbr">
<p>
Fin - Final Event<br/>
Pre - Prelim Event<br/>
Ind - Individual event<br/>
Rel - Relay event<br/>
PL - Place in Event<br/>
PT - Points Received<br/>
TR - Rank within My Times<br/>
<u>Date format</u><br/>dd/mm/y
</p>
</card>


<card id="qualify">
<p>
<b>Do I qualify</b><br/>
Is still under development.
In future, list of qualifying entries will appear, upon selecting one. It shall display what I qualify for 100%. 
That's why wait.
</p>
</card>

</wml>

<%}
rs.Close();
oConn.Close();
}else{
if(Request.QueryString("f") == "i")
{
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	var Athlete = Request.QueryString("a");
	rs = oConn.Execute("SELECT Athlete.Last, Athlete.First, Athlete.Initial, Athlete.Sex, Athlete.Birth, TEAM.TName FROM Athlete INNER JOIN TEAM ON Athlete.Team1 = TEAM.Team WHERE (((Athlete.Athlete)=" + Athlete + "));");

if(!rs.eof)
{
%>
<wml>
<card title="Ath Info">
<p align="center"><u>Athlete Info</u></p>
<p>Surname: <%=rs("Last")%></p>
<p>F.Name: <%=rs("First")%></p>
<p>Initials: <%=rs("Initial")%></p>
<p>Gender: <%if(rs("Sex")=="F"){%>Female<%}else{%>Male<%}%></p>
<p>Birth: <%=rs("Birth")%></p>
<p>Team: <%=rs("TName")%></p>

</card>
</wml>
<%
}else
{%>
<wml>

<card ontimer="time.asp?a=<%=Athlete%>" title="Error"><timer value="30"/>
<p align="center">
Error data not found, redirecting..
</p>
</card>
</wml>

<%}
}
else if(Request.QueryString("f") == "l")
{
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	
	var Athlete = Request.QueryString("a");
	var Course= Request.QueryString("c");
	rs = oConn.Execute("SELECT MEET.Meet, MEET.Start, Left([MName],18) AS TMName, MEET.Course FROM RESULT INNER JOIN MEET ON RESULT.MEET = MEET.Meet GROUP BY MEET.Meet, MEET.Start, Left([MName],18), MEET.Course, RESULT.ATHLETE HAVING (((RESULT.ATHLETE)=" + Athlete + ")) ORDER BY MEET.Start DESC;");

%>
<wml>
<card title="Meets">
<p>
<small>
<%if(!rs.eof){while(!rs.eof){%>
<a href="time.asp?a=<%=Athlete%>&amp;f=m&amp;m=<%=rs(0)%>"><%=rs(2)%></a> <%=rs(1)%> - <%=rs(3)%><br/><%rs.MoveNext();}}else{%>
No results found<%}%>
</small></p></card>
</wml>


<%}else{

if(Request.QueryString("f") == "m")
{
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	
	var Athlete = Request.QueryString("a");
	var Meet = Request.QueryString("m");
	rs = oConn.Execute("SELECT IIf([NT]=5,'DQ',IIf([NT]=2,'NS',[SCORE])) AS [time], RESULT.STROKE, RESULT.DISTANCE, RESULT.F_P, RESULT.I_R, (RESULT.POINTS/10) As points, RESULT.PLACE, RESULT.RANK FROM RESULT WHERE (((RESULT.ATHLETE)=" + Athlete +") AND ((RESULT.MEET)=" + Meet +")) ORDER BY RESULT.STROKE, RESULT.DISTANCE;");

	var strokes = new Array("","Freestyle","BackStroke","Breast","Butterfly","Medley");
	var menuit = new Array("","Freestyle (free)","BackStroke (back)","BreastStroke (brst)","Butterfly (fly)","Medaly (IM)");
%>
<wml>

<%
var results="";
var menu="<card id=\"stroke\"><p>" ;
var str = -1;
var cstroke = -1;
if(!rs.eof){while(!rs.eof){
str = rs(1).Value;%>
<%if(cstroke != str){
menu +="<a href=\"#s"+str+"\">"+menuit[str]+"</a><br/>";
results+="</small></p></card><card id=\"s"+rs(1)+"\"><p>";
results+="<b>"+strokes[rs(1)]+"</b><br/><small>";
cstroke = str}
results+= rs(2)+"m "+CTime(rs(0)) + "&nbsp;";
if(rs(3)=="F"){results+="Fin "}else{results+="Pre "}
if(rs(4)=="I"){results+="Ind "}else{results+="Rel "}
results+="Pl "+rs(6)+" &nbsp;&nbsp; PT "+rs(5)+" &nbsp;&nbsp;TR "+rs(7)+"<br/><br/>";
rs.MoveNext();}}else{menu+= "Sorry No Results"}%>

<%=menu%>
</p></card>
<card><p><small>
<%=results%>
</small></p></card>
</wml>
<%
}else{


if(Request.QueryString("f") == "t")
{
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	
	var Athlete = Request.QueryString("a");
	var Course = Request.QueryString("c")
	rs = oConn.Execute("SELECT IIf([NT]=5,'DQ',IIf([NT]=2,'NS',[SCORE])) AS [time], RESULT.STROKE, RESULT.DISTANCE, RESULT.F_P, RESULT.I_R, Day([Start]) & '/' & Month([Start]) & '/' & Right(Year([Start]),1) AS StateD, Left([MName],13) AS TMName FROM RESULT INNER JOIN MEET ON RESULT.MEET = MEET.Meet WHERE (((RESULT.ATHLETE)=" + Athlete +") AND ((RESULT.RANK)=1) AND ((RESULT.COURSE)='"+Course+"')) ORDER BY RESULT.STROKE, RESULT.DISTANCE;");

	var strokes = new Array("","Freestyle","BackStroke","Breast","Butterfly","Medley");
	var menuit = new Array("","Freestyle (free)","BackStroke (back)","BreastStroke (brst)","Butterfly (fly)","Medaly (IM)");
%>
<wml>


<%
var results="";
var menu="<card id=\"stroke\"><p>Top Times - " + Course + "<br />";
var str = -1;
var cstroke = -1;
if(!rs.eof){while(!rs.eof){
str = rs(1).Value;%>
<%if(cstroke != str){
menu +="<a href=\"#s"+str+"\">"+menuit[str]+"</a><br />";
results+="</small></p></card><card id=\"s"+rs(1)+"\"><p>";
results+="<b>"+strokes[rs(1)]+"</b><br /><small>";
cstroke = str}
results+= rs(2)+"m "+CTime(rs(0)) + "&nbsp;";
if(rs(3)=="F"){results+="Fin "}else{results+="Pre "}
if(rs(4)=="I"){results+="Ind "}else{results+="Rel "}
results+=rs(5)+" "+rs(6)+"<br/><br/>";
rs.MoveNext();}}else{menu+= "Sorry No Results"}%>

<%=menu%>
</p></card>
<card><p><small>
<%=results%>
</small></p></card>
</wml>


<%
}else{
if(Request.QueryString("f") == "a")
{
if(Request.QueryString("c").Count == 0 || Request.QueryString("s").Count == 0 || Request.QueryString("d").Count == 0)
{
%>
<wml>

<card title="All Times">
<p><u>Select Course</u><br/>
<select name="My_Course" title="Select Course">
<option value="S">Short 25m Pool</option>
<option value="L">Long 50m Pool</option>
</select><br/>

<u>Select Stroke</u><br/>
<select name="My_Stroke" title="Select Stroke">
<option value="1">Freestyle (free)</option>
<option value="2">BackStroke (back)</option>
<option value="3">BreastStroke (brst)</option>
<option value="4">Butterfly (fly)</option>
<option value="5">Medaly (IM)</option>
</select><br/>

<u>Select Distance</u><br/>
<select name="My_Distance" title="Select Distance">
<option value="25">25</option>
<option value="50">50</option>
<option value="100">100</option>
<option value="200">200</option>
<option value="400">400</option>
<option value="800">800</option>
<option value="1500">1500</option>
</select><br/>
<anchor><go href="time.asp?a=<%=Request.QueryString("a")%>&amp;f=a">
<postfield name="c" value="$(My_Course)" />
<postfield name="s" value="$(My_Stroke)" />
<postfield name="d" value="$(My_Distance)" />
</go>Get Times</anchor>
</p>
</card>
</wml>
<%
}else{
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	
	var Athlete = Request.QueryString("a");
	var Course = Request.QueryString("c")
	var Stroke = Request.QueryString("s")
	var Distance = Request.QueryString("d")
	rs = oConn.Execute("SELECT IIf([NT]=5,'DQ',IIf([NT]=2,'NS',[SCORE])) AS [time], RESULT.STROKE, RESULT.DISTANCE, RESULT.F_P, RESULT.I_R, Day([Start]) & '/' & Month([Start]) & '/' & Right(Year([Start]),1) AS StateD, Left([MName],13) AS TMName FROM RESULT INNER JOIN MEET ON RESULT.MEET = MEET.Meet WHERE (((RESULT.ATHLETE)=" + Athlete +") AND ((RESULT.STROKE)="+Stroke+") AND ((RESULT.DISTANCE)="+Distance+") AND ((RESULT.COURSE)='"+Course+"')) ORDER BY RESULT.STROKE, RESULT.DISTANCE,RESULT.RANK;");

	var strokes = new Array("","Free","Back","Breast","Butterfly","Medley");
%>
<wml>
<%
var results="";
results+="<b>"+strokes[Stroke]+" - " + Course + " - " + Distance + "<small>m</small></b><br/>";
results+="<small>";
if(!rs.eof){while(!rs.eof){
results+=rs(2)+"m "+CTime(rs(0)) + "&nbsp;";
if(rs(3)=="F"){results+="Fin "}else{results+="Pre "}
if(rs(4)=="I"){results+="Ind "}else{results+="Rel "}
results+=rs(5)+" "+rs(6)+"<br/><br/>";
rs.MoveNext();}}else{results+= "Sorry No Results"}%>

<card><p>
<%=results%>
</small>
</p></card>
</wml>


<%
}
}else{%>

<wml>
<card>
<p>Error that is not an option</p>
</card>
</wml><%}}}}}%>