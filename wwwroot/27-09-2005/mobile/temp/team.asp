<%@ Language=JavaScript %>
<% Response.ContentType = "text/vnd.wap.wml" %>
<%Response.Expires = 2880%>

<?xml version="1.0"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"> 

<%
//test for alpha menu chrater
if(Request.QueryString("Alpha").Count == 0)
{%>
<wml>
<card>
Enter Club First Letter<br />
<fieldset>
<input name="My_Alpha" format="A" emptyok="false" maxlength="1"/><br/>
</fieldset>
<anchor><go method="get" href="team.asp">
<postfield name="p" value="WP" />
<postfield name="Alpha" value="$(My_Alpha)" />
</go>Get club list</anchor>
</card></wml>
<%}else{

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
Select Club - '<%=Alpha%>'<br/>
<p nowrap>
<%if(!rs.eof){while(!rs.eof){%>
<a href="ath.asp?t=<%=rs(0)%>"><%=rs(1)%></a><br/><%rs.MoveNext();
}}else{
%>Sorry no clubs<%
}rs.Close();oConn.Close();%>
</p>
</card>
</wml><%}%>