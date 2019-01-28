<%@ Language=JavaScript %>
<% Response.ContentType = "text/vnd.wap.wml" %>
<%Response.Expires = 2880%>
<?xml version="1.0"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"> 

<%
//test for alpha menu chrater
if(Request.QueryString("t").count != 0)
{
if(Request.QueryString("Alpha").Count == 0)
{%>
<wml>
<card newcontext="true">
<setvar name="My_Alpha" value="" />
Enter Athlete First Letter<br />
<fieldset>
<input name="My_Alpha" format="A" emptyok="false" maxlength="1"/>
</fieldset><br/>
<anchor>Get list<go method="get" href="ath.asp">
<postfield name="t" value="<%=Request.QueryString("t")%>" />
<postfield name="Alpha" value="$(My_Alpha)" />
<setvar name="My_Alpha" value="" />
</go>
</anchor>
</card></wml>
<%}else{
	var Team = Request.QueryString("t");
	var Alpha = Request.QueryString("Alpha");
	
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	rs = oConn.Execute("SELECT Athlete.Athlete, Left(Athlete.First,1) As initial, Athlete.Last, Athlete.Sex,  Year([Birth]) AS bryear FROM Athlete WHERE (((Athlete.Team1)=" + Team + ") AND (UCase(Left(Athlete.Last,1)) ='" + Alpha +"') AND Athlete.Group = 'A') ORDER BY Athlete.First")
	
	%>


<wml>
<card id="card1">
Select Athlete - <%=Alpha%><br/>
<p nowrap>
<%if(!rs.eof){while(!rs.eof){%>
<a href="time.asp?a=<%=rs(0)%>"><%=rs(1)%>.<%=rs(2)%> - <%=rs(3)%> - <%=rs(4)%></a><br/><%rs.MoveNext();
}}else{
%>Sorry no Athlete's<%
}rs.Close();oConn.Close();%>
</p>

</card></wml><%}}else{%>error<%}%>