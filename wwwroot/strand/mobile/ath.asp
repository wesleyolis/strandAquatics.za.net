<%@ Language=JavaScript %>
<%Response.ContentType = "text/vnd.wap.wml"
%><?xml version="1.0" encoding="iso-8859-1" ?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"> 


<%Response.Expires = 2880%>
<%
//test for alpha menu chrater
if(Request.QueryString("t").count != 0)
{
if(Request.QueryString("L").count == 0)
{%>
<wml>


<card title="Athlete Surename">
<p>Athlete Sure Letter<br/>
<table columns="7" title="Athlete F.Letter">
<tr>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=A">A</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=B">B</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=C">C</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=D">D</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=E">E</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=F">F</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=G">G</a></td>
</tr>
<tr>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=H">H</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=I">I</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=J">J</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=K">K</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=L">L</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=M">M</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=N">N</a></td>
</tr>
<tr>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=O">O</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=P">P</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=Q">Q</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=R">R</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=S">S</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=T">T</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=U">U</a></td>
</tr>
<tr>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=V">V</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=W">W</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=X">X</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=Y">Y</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=Z">Z</a></td>
<td></td>
<td></td>
</tr>
</table>
</p>
</card>

</wml>
<%}else
{
if(Request.QueryString("F").count == 0)
{%>
<wml>


<card title="Athlete Surename">
<p>Athlete Name Letter<br/>
<table columns="7" title="Athlete F.Letter">
<tr>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=A">A</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=B">B</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=C">C</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=D">D</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=E">E</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=F">F</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=G">G</a></td>
</tr>
<tr>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=H">H</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=I">I</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=J">J</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=K">K</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=L">L</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=M">M</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=N">N</a></td>
</tr>
<tr>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=O">O</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=P">P</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=Q">Q</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=R">R</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=S">S</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=T">T</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=U">U</a></td>
</tr>
<tr>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=V">V</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=W">W</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=X">X</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=Y">Y</a></td>
<td><a href="ath.asp?t=<%=Request.QueryString("t")%>&amp;L=<%=Request.QueryString("L")%>&amp;F=Z">Z</a></td>
<td></td>
<td></td>
</tr>
</table>
</p>
</card>

</wml>
<%}else{

	var Team = Request.QueryString("t");
	var Last = Request.QueryString("L");
	var First = Request.QueryString("F");
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("../Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	rs = oConn.Execute("SELECT Athlete.Athlete, Left(Athlete.First,1) As initial, Athlete.Last, Athlete.Sex,  Year([Birth]) AS bryear FROM Athlete WHERE (((Athlete.Team1)=" + Team + ") AND (UCase(Left(Athlete.Last,1)) = UCase('" + Last +"')) AND (UCase(Left(Athlete.First,1)) = UCase('" + First +"')) AND Athlete.Group = 'A') ORDER BY Athlete.Last, Year([Birth])")
	
if(!rs.eof){
	%>
<wml>


<card title="Select Athlete">
<p>Select Athlete - <%=Last%> - <%=First%><br/>
<%while(!rs.eof){%>
<a href="time.asp?a=<%=rs(0)%>"><%=rs(1)%>.<small><%=rs(2)%> - <%=rs(3)%> - <%=rs(4)%></small></a><br/><%rs.MoveNext();
}}else{
%>
<wml>

<card ontimer="ath.asp?t=<%=Team%>" title="Non Found"><timer value="30"/>
<p align="center">
Sorry no Athlete's Found, redirecting..
</p>
</card>
</wml>
<%
}rs.Close();oConn.Close();%>
</p>
</card>

</wml><%}}}else{
%>

<wml>

<card ontimer="index.wml" title="Not Found"><timer value="30"/>
<p align="center">
Athlete not found, redirecting..
</p>
</card>
</wml>


<%}%>