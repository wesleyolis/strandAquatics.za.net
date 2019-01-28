<%@ Language=JavaScript %>
<!--METADATA TYPE="typelib" 
uuid="00000206-0000-0010-8000-00AA006D2EA4" -->
<%

Response.CacheControl = "No-Cache";
Response.Expires = -1;%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<%


var oConn;	
	var rs,rst,exclude_meet = "",SQL="",ath = -1, course = -1, dis = -1,stroke = -1,rank = -1;		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	
		rst = oConn.Execute("SELECT MEET.Meet FROM MEET WHERE (((InStr(1,[MName],'Best Times'))<>0));");
	if(!rst.eof)
	{
		exclude_meet = "((RESULT.MEET)<>" + rst.Fields("meet") + ")";
		rst.MoveNext();
	}
	while(!rst.eof)
	{
	exclude_meet += " AND ((RESULT.MEET)<>" + rst.Fields("meet") + ")";
	rst.MoveNext();
	}


SQL +="SELECT RESULT.TRANK, RESULT.ATHLETE, RESULT.COURSE, RESULT.STROKE, RESULT.DISTANCE, RESULT.SCORE";
SQL +="FROM RESULT";
SQL +="WHERE (((RESULT.ATHLETE)<>0) AND ((RESULT.SCORE)>0) AND (" + exclude_meet +"))";
SQL +="ORDER BY RESULT.ATHLETE, RESULT.COURSE, RESULT.STROKE, RESULT.DISTANCE, RESULT.SCORE;";


'--- Open a static cursor on the Survey table, and add record

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open( SQL, oConn,adOpenStatic, adLockOptimistic, "&H0001");


While(!rs.EOF)
{
	if(rs.Fields("DISTANCE").Value != dis && rs.Fields("STROKE").Value != stroke && rs.Fields("COURSE").Value != course && rs.Fields("ATHLETE").Value != ath)
	{rank = 1;
	dis = rs.Fields("DISTANCE").Value;
	stroke = rs.Fields("STROKE").Value;
	course = rs.Fields("COURSE").Value;
	ath = rs.Fields("ATHLETE").Value;
	}
	rs("TRANK") = rank;
	rs.Update();
	rs.MoveNext();
}

rs.Close();


%>


</head>

<body>

</body>

</html>
