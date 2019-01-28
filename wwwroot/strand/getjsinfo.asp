<%@ Language=JavaScript %>
<%
Response.CacheControl = "Public";
Response.Expires = 1440;
Response.Buffer = false;
%>
<%
var oConn, rs;
function connect()
{
		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
}
	 
var option  = Request.QueryString("data");	 
var key  = Request.QueryString("Key");	 
if (option == "ath_time")
{
if(key > 0 && key <= 65536)
{
connect();

rs = oConn.Execute("SELECT RESULT.RESULT, MEET.Meet, RESULT.MTEVENT, RESULT.COURSE, RESULT.I_R, RESULT.F_P, RESULT.DISTANCE, RESULT.STROKE, RESULT.NT, RESULT.SCORE, RESULT.AGE, RESULT.PLACE, RESULT.POINTS, MEET.MName FROM RESULT INNER JOIN MEET ON RESULT.MEET = MEET.Meet WHERE (((RESULT.ATHLETE)=" + key + "));");
Response.write("var coloms=14,start=3,result=0,meet=1,mevent=2,course=3,I_R=4,F_P=5,dis=6,stroke=7,nt=8,score=9,age=10,place=11,points=13,mname=14;\n");
Response.write("var ath_times_widths=new Array(0,0,0,40,40,70,50,50,0,60,40,50,50,150);\n");
Response.write("var ath_times_names=new Array(0,0,0,'Course','Indi/Relay','Final/Pre/semi','Distance','Stroke','Tinfo','Time','Age','Place','Points','Meet');\n");
Response.write("var ath_times=new Array(");
while(!rs.eof)
{
Response.write(rs(0) + ",");
Response.write(rs(1) + ",");
Response.write(rs(2) + ",");
Response.write("'" + rs(3) + "',");
Response.write("'" + rs(4) + "',");
Response.write("'" + rs(5) + "',");
Response.write("" + rs(6) + ",");
Response.write(rs(7) + ",");
Response.write(rs(8) + ",");
Response.write(rs(9) + ",");
Response.write(rs(10) + ",");
Response.write(rs(11) + ",");
Response.write(rs(12) + ",");
Response.write("'" + rs(13)  + "',");
rs.MoveNext();
}
Response.Write("'');");
rs.close();
oConn.close();
}}
%>