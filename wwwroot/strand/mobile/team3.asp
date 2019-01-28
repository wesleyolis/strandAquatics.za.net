<%@ Language=JavaScript %>
<%Response.ContentType = "text/vnd.wap.wml" %>
<%Response.Expires = 2880%>

<?xml version="1.0"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"> 
<%
//test for alpha menu chrater
if(Request.Form("Alpha").Count == 0)
{%>
<wml>
<card>
<do type="ACCEPT" label="Get Club List">
<go method="post" href="team3.asp" accept-charset="UTF-8" >
<postfield name="p" value="WP"/>
<postfield name="Alpha" value="$My_Alpha"/>
</go>
</do>

<p>
Enter Club First Letter<br />
<fieldset>
<input name="My_Alpha" format="A" emptyok="false" maxlength="1"/><br/>
</fieldset>

</p>
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
	var LSC = Request.Form("P");
	var Alpha = Request.Form("Alpha");
	rs = oConn.Execute("SELECT TEAM.Team, TEAM.TName FROM TEAM WHERE (((TEAM.LSC)='" + LSC + "') AND UCase(Left(TEAM.TName,1)) =UCase('" + Alpha + "')) ORDER BY TEAM.TName;")
	}
	%>


<wml>
<card>
<p>
Select Club - <%=Alpha%><br/>

<%if(!rs.eof){while(!rs.eof){%>
<a href="ath.asp?t=<%=rs(0)%>"><%=rs(1)%></a><br/><%rs.MoveNext();
}}else{
%>Sorry no clubs<%
}rs.Close();oConn.Close();%>
</p>
</card>
</wml><%}%>