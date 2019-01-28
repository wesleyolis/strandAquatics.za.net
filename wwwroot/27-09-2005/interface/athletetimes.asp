<%@ Language=JavaScript %>
<%
Response.Buffer=true;
Response.CacheControl = "Public";
Response.Expires = 10;

var strokes = new Array("","Freestyle","Back","Breast","Butterfly","Medley");
	var oConn;	
	var rs;		
	var filePath;	
	 if (Request.QueryString("ath").count == 0)
	 {%>-1|
	 <%
	 }
	 else
	 {
	 filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	var Athlete = Request.QueryString("ath");
	//WHERE (((RESULT.ATHLETE)=" + Athlete + "))
	rs = oConn.Execute("SELECT RESULT.COURSE, RESULT.STROKE, RESULT.DISTANCE, RESULT.NT, RESULT.SCORE, RESULT.F_P, RESULT.I_R, RESULT.AGE, RESULT.MEET, RESULT.MTEVENT FROM RESULT INNER JOIN MEET ON RESULT.MEET = MEET.Meet WHERE (((RESULT.ATHLETE)=" + Athlete + ")) ORDER BY RESULT.COURSE, RESULT.STROKE, RESULT.DISTANCE, RESULT.NT, RESULT.SCORE;");
%>
['timeformat.js']>>><%
while(!rs.eof)
{
for(i=0;i<10;i++){
%><%=rs(i)%><%if(i!=9){%>|<%}}
%>><%
rs.MoveNext();
}
rs.close();
oConn.close();%><%
}
%>