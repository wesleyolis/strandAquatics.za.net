<%@ Language=JavaScript %>
<% Response.ContentType = "text/vnd.wap.wml" %>
<%Response.Expires = 2880%>
<%
function CTime (Score)
{
if(Score != "DQ" & Score != "NS")
{

if(Score<6000)
{t = "00:" + (Score/100);
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
else{t = Score}//dq
return t;}



%>


<?xml version="1.0"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"> 

<%
//test for alpha menu chrater
if(Request.QueryString("f").Count == 0)
{
%>
<wml>

<card id="menu">
<u>Athlete Menu</u><br />
<a href="time.asp?a=<%=Request.QueryString("a")%>&f=l">Latest Meet Times</a><br />
<a href="#course">Top Times</a><br />
<a href="time.asp?a=<%=Request.QueryString("a")%>&f=a">All Times</a><br />
<a href="#qualify">Do I qualify</a><br />
<a href="#abbr"><b>Abbriviations</b></a><br />
</card>

<card id="course">
<u>Select Course</u><br />
<select name="My_Course">
<option value="S">Short 25m Pool<onevent type="onpick"><go method="get" href="time.asp?a=<%=Request.QueryString("a")%>&f=t">
<postfield name="c" value="$(My_Course)" /></go></onevent></option>
<option value="L">Long 50m Pool<onevent type="onpick"><go method="get" href="time.asp?a=<%=Request.QueryString("a")%>&f=t">
<postfield name="c" value="$(My_Course)" /></go></onevent></option>
</select>
</card>




<card id="abbr">
Fin - Final Event<br/>
Pre - Prelim Event<br/>
Ind - Individual event<br/>
Rel - Realy event<br/>
PL - Place in Event<br/>
PT - Points Received<br/>
TR - Rank within My Times<br/>
<u>Date format</u><br/>dd/mm/y
</card>





<card id="qualify">
<b>Do I qualify</b><br/>
Is still under development.
In future, list of qualifying entries will appear, apon selecting one. It shall display what I qualify for 100%. Thats why wait.
</card>
</wml>

<%}else{
if(Request.QueryString("f") == "l")
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
<card><small>
<%if(!rs.eof){while(!rs.eof){%>
<a href="time.asp?a=<%=Athlete%>&f=m&m=<%=rs(0)%>"><%=rs(2)%></a> <%=rs(1)%> - <%=rs(3)%><br/><%rs.MoveNext();}}else{%>
No results found<%}%>
</small></card>
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
var menu="<card id=\"stroke\">" ;
var str = -1;
var cstroke = -1;
if(!rs.eof){while(!rs.eof){
str = rs(1).Value;%>
<%if(cstroke != str){
menu +="<a href=\"#s"+str+"\">"+menuit[str]+"</a><br />";
results+="</small></card><card id=\"s"+rs(1)+"\">";
results+="<b>"+strokes[rs(1)]+"</b><br /><small>";
cstroke = str}
results+= rs(2)+"m "+CTime(rs(0)) + "&nbsp;";
if(rs(3)=="F"){results+="Fin "}else{results+="Pre "}
if(rs(4)=="I"){results+="Ind "}else{results+="Rel "}
results+="Pl "+rs(6)+" &nbsp;&nbsp; PT "+rs(5)+" &nbsp;&nbsp;TR "+rs(7)+"<br/><br/>";
rs.MoveNext();}}else{menu+= "Sorry No Results"}%>

<%=menu%>
</card>
<card><small>
<%=results%>
</small></card>







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
var menu="<card id=\"stroke\">Top Times - " + Course + "<br />";
var str = -1;
var cstroke = -1;
if(!rs.eof){while(!rs.eof){
str = rs(1).Value;%>
<%if(cstroke != str){
menu +="<a href=\"#s"+str+"\">"+menuit[str]+"</a><br />";
results+="</small></card><card id=\"s"+rs(1)+"\">";
results+="<b>"+strokes[rs(1)]+"</b><br /><small>";
cstroke = str}
results+= rs(2)+"m "+CTime(rs(0)) + "&nbsp;";
if(rs(3)=="F"){results+="Fin "}else{results+="Pre "}
if(rs(4)=="I"){results+="Ind "}else{results+="Rel "}
results+=rs(5)+" "+rs(6)+"<br/><br/>";
rs.MoveNext();}}else{menu+= "Sorry No Results"}%>

<%=menu%>
</card>
<card><small>
<%=results%>
</small></card>
</wml>


<%
}else{
if(Request.QueryString("f") == "a")
{
if(Request.QueryString("c").Count == 0 || Request.QueryString("s").Count == 0 || Request.QueryString("d").Count == 0)
{
%>
<wml>

<card>
<u>Select Course</u><br />
<select name="My_Course">
<option value="S">Short 25m Pool</option>
<option value="L">Long 50m Pool</option>
</select>

<u>Select Stroke</u><br />
<select name="My_Stroke">
<option value="1">Freestyle (free)</option>
<option value="2">BackStroke (back)</option>
<option value="3">BreastStroke (brst)</option>
<option value="4">Butterfly (fly)</option>
<option value="5">Medaly (IM)</option>
</select><br />

<u>Select Distance</u><br />
<select name="My_Distance">
<option value="25">25</option>
<option value="50">50</option>
<option value="100">100</option>
<option value="200">200</option>
<option value="400">400</option>
<option value="800">800</option>
<option value="1500">1500</option>
</select>
<anchor>Get Times<go href="time.asp?a=<%=Request.QueryString("a")%>&f=a">
<postfield name="c" value="$(My_Course)" />
<postfield name="s" value="$(My_Stroke)" />
<postfield name="d" value="$(My_Distance)" />
</go>
</anchor>
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
<%=results%>
</small>
</card>
</wml>


<%
}
}else{





	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	if(Request.QueryString("t").count == 0)
	{
	var LSC = Request.QueryString("P");
	var Alpha = Request.QueryString("Alpha");
	rs = oConn.Execute("SELECT TEAM.Team, TEAM.TName FROM TEAM WHERE (((TEAM.LSC)='" + LSC + "') AND UCase(Left(TEAM.TName,1)) ='" + Alpha + "') ORDER BY TEAM.TName;")
	}
	%>


<wml>
<card id="card1">
Select Club - <%=Alpha%><br/>
<p nowrap>
<%if(!rs.eof){while(!rs.eof){%>
<a href="ath.asp?t=<%=rs(0)%>"><%=rs(1)%></a><br/><%rs.MoveNext();
}}else{
%>Sorry no clubs<%
}rs.Close();oConn.Close();%>
</p>
</card>
</wml><%}}}}}%>