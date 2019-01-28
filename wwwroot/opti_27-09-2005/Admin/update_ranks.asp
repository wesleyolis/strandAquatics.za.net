<%@ Language=JavaScript %>
<!--METADATA TYPE="typelib" 
uuid="00000206-0000-0010-8000-00AA006D2EA4" -->
<%

Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'DB Ranking Season Times';\n");Response.write("/--></script>\n");


var oConn;	
	var rs,rst,exclude_meet = "",include_meet="",SQL="",ath = -1, course = -1, dis = -1,stroke = -1,rank = -1,time = -1;		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	
		rst = oConn.Execute("SELECT MEET.Meet FROM MEET WHERE (((InStr(1,[MName],'Best Times'))<>0));");
	if(!rst.eof)
	{
		exclude_meet = "((RESULT.MEET)<>" + rst.Fields("meet") + ")";
		include_meet = "((RESULT.MEET)=" + rst.Fields("meet") + ")";

		rst.MoveNext();
	}
	while(!rst.eof)
	{
	exclude_meet += " AND ((RESULT.MEET)<>" + rst.Fields("meet") + ")";
	include_meet += " OR ((RESULT.MEET)=" + rst.Fields("meet") + ")";
	rst.MoveNext();
	}
	
	rst.Close();
	rst = null;
	
Response.Write("UPDATE RESULT SET RESULT.TRANK = 0 WHERE (" + include_meet +");");

oConn.Execute("UPDATE RESULT SET RESULT.TRANK = 0 WHERE (" + include_meet +");");

Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'DB Ranking Season Times 1/2';\n");Response.write("/--></script>\n");


SQL +="SELECT RESULT.TRANK, RESULT.ATHLETE, RESULT.COURSE, RESULT.STROKE, RESULT.DISTANCE, RESULT.SCORE";
SQL +=" FROM RESULT";
SQL +=" WHERE ((((RESULT.PLACE)<>0) AND (RESULT.ATHLETE)<>0) AND ((RESULT.SCORE)>0) AND (" + exclude_meet +"))";
SQL +=" ORDER BY RESULT.ATHLETE, RESULT.COURSE, RESULT.STROKE, RESULT.DISTANCE, RESULT.SCORE;";

Response.Write("<br><br>" + SQL);

rs = Server.CreateObject("ADODB.Recordset");
rs.Open( SQL, oConn,adOpenStatic, adLockOptimistic, "&H0001");


while(!rs.eof)
{

	if(rs.Fields("DISTANCE").Value != dis || rs.Fields("STROKE").Value != stroke || rs.Fields("COURSE").Value != course || rs.Fields("ATHLETE").Value != ath)
	{rank = 1;
	dis = rs.Fields("DISTANCE").Value;
	stroke = rs.Fields("STROKE").Value;
	course = rs.Fields("COURSE").Value;
	ath = rs.Fields("ATHLETE").Value;
	}
	if(time != rs.Fields("SCORE").Value)
	{
	time = rs.Fields("SCORE").Value;
	rs("TRANK") = rank;
	rank++;
	}
	else
	{
	rs("TRANK") = 0;
	}
	
	rs.Update();
	
	rs.MoveNext();
	
}

rs.Close();
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'DB Ranking Season Times Done';\n");Response.write("/--></script>\n");


%>


