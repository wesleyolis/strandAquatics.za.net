<%@ Language=JavaScript %>
<%Response.ContentType = "text/vnd.wap.wml"
%><?xml version="1.0" encoding="iso-8859-1" ?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">


<!--METADATA TYPE="typelib" 
uuid="00000206-0000-0010-8000-00AA006D2EA4" -->

<%Response.Expires = 2880%>
<%

function Age(Lo,Hi)
{if(Lo <  Hi)
{
if(Lo > 0 & Hi < 99)
{
age = Lo + "-" + Hi + "yrs";
}else
{
if(Hi == 99 & Lo >0)
{
age = Lo + "&amp;Ovr";
}
else
{
if(Lo == 0 & Hi < 99)
{
age = Hi + "&amp;Und";
}
else
{
if(Lo == 0 & Hi == 99)
{
age = "Open";
}}}}}else{age = Hi + "yrs";}

return age;
}

function CTime (Score)
{
if(Score ==0)
{return "- -"}
else
{
s2=""; m2="";h2="";
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

return t;}}

function posit(pos,NT)
{if(pos == 0){
if(NT == 2){return "NS";}
else if(NT==5){return "DQ";}
}
return pos;
}

var stroke = new Array("","Free","Back","Brst","Fly","IM");

//test for alpha menu chrater
if(Request.QueryString("m").Count == 0 & Request.QueryString("e").Count == 0)
{
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	rs = Server.CreateObject("ADODB.Recordset");
	rs.Open ("SELECT MEET.Meet, MEET.MName, MEET.Start, MEET.Course FROM MEET WHERE ((MEET.Start) < Date()) ORDER BY MEET.Start DESC;", oConn, adOpenStatic);
	rs.PageSize = 10;
	if(rs.PageCount !=0)
	{
	rs.AbsolutePage = 1;
	}

%>
<wml>
<card title="Meets List">
<p>Latest 5 Meets</p>
<%if(rs.eof){%><p>No Meets found</p><%}else
for(j=0; (j < rs.PageSize) && (!rs.eof); j++, rs.MoveNext())
{%>
<p><small><a href="meet.asp?m=<%=rs(0)%>&amp;p=1"><%=rs(1)%><br/>Start: <%=rs(2)%><br/>Course: <%=rs(3)%></a></small><br/></p>
<%
}
%>
</card>

</wml>

<%}else if(Request.QueryString("e").Count == 0)
{
	
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	
	var Meet = Request.QueryString("m");
	var Page = Request.QueryString("p");

	rs = Server.CreateObject("ADODB.Recordset");
	rs.Open ("SELECT MTEVENT.MtEvent, [MtEv] & [MtEvX] AS Event, MTEVENT.Sex, (MTEVENT.Lo_Hi Mod 100) As Hi,((MTEVENT.Lo_Hi - [Hi])/100) as low, MTEVENT.Distance, MTEVENT.Stroke, MTEVENT.I_R FROM MTEVENT WHERE (((MTEVENT.Meet)=" + Meet + ") And (MTEVENT.I_R='I')) ORDER BY [MtEv],[MtEvX];", oConn, adOpenStatic);
	rs.PageSize = 10;
	
	if((Page>0) & (Page < rs.PageCount) & rs.PageCount !=0)
	{
	rs.AbsolutePage = Page;
	}
	else
	{
	//error redict
	}
	%>


<wml>


<card title="Meet Events">
<%if(rs.eof){%><p>No Events found</p><%}else{if(rs.PageCount > 1){%>
<p align="center"><a href="#page">Page <%=Page%>/<%=rs.PageCount%></a></p>
<%}for(j=0; (j < rs.PageSize) & (!rs.eof); j++, rs.MoveNext())
{%>
<p><a href="meet.asp?e=<%=rs(0)%>&amp;t=P&amp;p=1"><%=rs(1)%>&nbsp;&nbsp;<%if(rs(2)=="F"){%>Fem<%}else if(rs(2) == "M"){%>Mal<%}else{%>Mix<%}%>&nbsp;&nbsp;<%=Age(rs(4),rs(3))%><br/>&nbsp;&nbsp;&nbsp;&nbsp;<%=rs(5)%>m <%=stroke[rs(6)]%>&nbsp;<%if(rs(7)=="I"){%>Ind<%}else{%>Rel<%}%></a><br/></p>
<%
}if(rs.PageCount > 1){
%>
<p align="center"><a href="#page">Page <%=Page%>/<%=rs.PageCount%></a></p>
<%}}%>
</card>

<%if(rs.PageCount > 1){%>
<card id="page" title="Event page">
<p><%for(j=1;j<=rs.PageCount;j++){if(j!=Page){%><a href="meet.asp?m=<%=Meet%>&amp;p=<%=j%>"><%=j%></a> 
<%}else{%>
<%=j%>&nbsp;
<%}}%></p>
</card>
<%}%>

</wml>
<%}
else
{

	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	
	var Event = Request.QueryString("e");
	var Type = Request.QueryString("t");
	var Page = Request.QueryString("p");

	rs = Server.CreateObject("ADODB.Recordset");
	rs.Open ("SELECT Left([First],1) AS Fir, Athlete.Last, Athlete.Sex, RESULT.AGE, RESULT.NT, RESULT.SCORE, RESULT.PLACE FROM Athlete INNER JOIN RESULT ON Athlete.Athlete = RESULT.ATHLETE WHERE (((RESULT.MTEVENT)="+Event+") And ((RESULT.F_P)='"+ Type +"')) ORDER BY RESULT.NT, RESULT.PLACE;", oConn, adOpenStatic);
	rs.PageSize = 15;
	
	if((Page>0) & (Page <= rs.PageCount) & rs.PageCount !=0)
	{
	rs.AbsolutePage = Page;
	}
	else
	{
	//error redict
	}
	pCount = rs.PageCount;
	%>


<wml>


<card title="Event Results - <%=Type%>">
<%if(rs.eof){%><p>No Results Found</p><%}else{
if(pCount > 1){%>
<p align="center"><a href="#page">Change Page <%=Page%>/<%=rs.PageCount%></a></p>
<%}%><%for(j=1; (j <= rs.PageSize) & (!rs.eof);j++, rs.MoveNext())
{%>
<p><%=posit(rs(6),rs(4))%>&nbsp;<small><%=rs(0)%>.<%=rs(1)%></small><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=CTime(rs(5))%>&nbsp;&nbsp;<%=rs(2)%>&nbsp;<%=rs(3)%><br/></p>
<%
}%><%if(pCount > 1){
%>
<p align="center"><a href="#page">Page <%=Page%>/<%=rs.PageCount%></a></p>
<%}}%>
</card>

<%if(pCount > 1){%>
<card id="page" title="Results Page">
<p><%for(j=1;j<=pCount;j++){if(j!=Page){%><a href="meet.asp?e=<%=Event%>&amp;t=<%=Type%>&amp;p=<%=j%>"><%=j%></a> 
<%}else{%>
<%=j%>&nbsp;
<%}}%></p>
<p><a href="meet.asp?e=<%=Event%>&amp;t=F&amp;p=1">Final (F)</a><br/>
<a href="meet.asp?e=<%=Event%>&amp;t=I&amp;p=1">Semi (I)</a><br/>
<a href="meet.asp?e=<%=Event%>&amp;t=P&amp;p=1">Prelims (P)</a></p>
</card>
<%}%></wml><%}%>