<%@ Language=JavaScript %>
<%Response.ContentType = "text/vnd.wap.wml"
%><?xml version="1.0" encoding="iso-8859-1" ?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">


<%Response.Expires = 2880%>
<%
//test for alpha menu chrater
if(Request.QueryString("Alpha").Count == 0)
{%>
<wml>


<card title="Club Letter">
<p>
Clubs First Letter<br/>
<table columns="7" title="Club Letter">
<tr>
<td><a href="team.asp?Alpha=A">A</a></td>
<td><a href="team.asp?Alpha=B">B</a></td>
<td><a href="team.asp?Alpha=C">C</a></td>
<td><a href="team.asp?Alpha=D">D</a></td>
<td><a href="team.asp?Alpha=E">E</a></td>
<td><a href="team.asp?Alpha=F">F</a></td>
<td><a href="team.asp?Alpha=G">G</a></td>
</tr>
<tr>
<td><a href="team.asp?Alpha=H">H</a></td>
<td><a href="team.asp?Alpha=I">I</a></td>
<td><a href="team.asp?Alpha=J">J</a></td>
<td><a href="team.asp?Alpha=K">K</a></td>
<td><a href="team.asp?Alpha=L">L</a></td>
<td><a href="team.asp?Alpha=M">M</a></td>
<td><a href="team.asp?Alpha=N">N</a></td>
</tr>
<tr>
<td><a href="team.asp?Alpha=O">O</a></td>
<td><a href="team.asp?Alpha=P">P</a></td>
<td><a href="team.asp?Alpha=Q">Q</a></td>
<td><a href="team.asp?Alpha=R">R</a></td>
<td><a href="team.asp?Alpha=S">S</a></td>
<td><a href="team.asp?Alpha=T">T</a></td>
<td><a href="team.asp?Alpha=U">U</a></td>
</tr>
<tr>
<td><a href="team.asp?Alpha=V">V</a></td>
<td><a href="team.asp?Alpha=W">W</a></td>
<td><a href="team.asp?Alpha=X">X</a></td>
<td><a href="team.asp?Alpha=Y">Y</a></td>
<td><a href="team.asp?Alpha=Z">Z</a></td>
<td></td>
<td></td>
</tr>
</table>
</p>
</card>

</wml>

<%}else{

	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	if(Request.QueryString("t").count == 0)
	{
	var LSC = "WP"; // QueryString("P");
	var Alpha = Request.QueryString("Alpha");
	rs = oConn.Execute("SELECT TEAM.Team, TEAM.TName FROM TEAM WHERE (((TEAM.LSC)='" + LSC + "') AND UCase(Left(TEAM.TName,1)) =UCase('" + Alpha + "')) ORDER BY TEAM.TName;")
	}
	%>


<wml>


<card title="Select Club">
<p>Select Club - <%=Alpha%><br/><%
if(!rs.eof){while(!rs.eof){
%>
<a href="ath.asp?t=<%=rs(0)%>"><%=rs(1)%></a><br/><%rs.MoveNext();
}}else{
%>Sorry no clubs<%
}rs.Close();oConn.Close();%>
</p>
</card>
</wml><%}%>