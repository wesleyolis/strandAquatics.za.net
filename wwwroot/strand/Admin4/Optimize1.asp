<%@ Language=JavaScript %>
<%
Response.CacheControl = "no-store";
Response.Expires = 0;
Response.Buffer = false;
if( Request.QueryString("DBName").count > 0 & Request.QueryString("Version").count > 0  & Request.QueryString("TMV").count > 0 )
{
	var pass = new Array( );
	var conn,oConn, Engine, fs;	
	var rs;		
	var filePath, OpenString;	
	var database = Request.QueryString("DBName");
	filePath = "C:/Inetpub/wwwroot/Swim Database/extration/" + database + ".mdb";
	if(Request.QueryString("TMV"))
	{
	OpenString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath;
	}
	else
	{
		conn = false;
		for(i=0;i<5 & ( conn == false);i++)
		{
			try
			{
				OpenString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath + ";Jet OLEDB:Database   +  ";";
				//reset the password on the database to nothing to be compatible with rest of site.
				conn = true;
				Response.Write("Connected to Database" + "<br>");
			}
			catch(System)
			{
				Response.Write("Error connecting to Database, inform wesley please");
			}
		}
		
		
	}
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open(OpenString);
	
	oConn.Execute ("DROP INDEX athWMgroup ON Athlete;");
	oConn.Execute ("ALTER TABLE Athlete DROP COLUMN Team2, Team3, Batch, Citizen, Pref, WMGroup, WMSubGr;");
	oConn.Execute ("UPDATE Athlete SET Athlete.[Last] = UCase(Left([Last],1)) & LCase(Mid([Last],2)), Athlete.[First] = UCase(Left([First],1)) & LCase(Mid([First],2)), Athlete.Initial = UCase(Left([Initial],1));"); 
	oConn.Execute ("DROP INDEX TEAM ON RELAY;");
	oConn.Execute ("DROP INDEX ATH1 ON RELAY;");
	oConn.Execute ("DROP INDEX ATH2 ON RELAY;");
	oConn.Execute ("DROP INDEX ATH3 ON RELAY;");
	oConn.Execute ("DROP INDEX ATH4 ON RELAY;");
	oConn.Execute ("ALTER TABLE RELAY DROP COLUMN [ATH(5)],[ATH(6)], [ATH(7)], [ATH(8)], RELAYAGE;");
	oConn.Execute ("DROP INDEX Athlete ON RESULT;");
	oConn.Execute ("DROP INDEX Meet ON RESULT;");
	oConn.Execute ("DROP INDEX Fast ON RESULT;");
	oConn.Execute ("DROP INDEX FastRelay ON RESULT;");
	oConn.Execute ("ALTER TABLE RESULT DROP COLUMN SPLIT, EX, ORIGIN, NT, MISC, TRANK;");
	oConn.Execute ("ALTER TABLE TEAM DROP COLUMN MailTo, TAddr, TCity, TState, TZip, TCntry, [Day], Eve, Fax, TType, Regn, Reg, MinAge, NumAth, EMail, TM2000, TDivision;");
	oConn.Execute ("DROP TABLE  ATHRECR, AthInfo, Attendance, CUSTOMAGEGROUP, COACHES, CUSTOMRPTS, Energy, ENTRY, ESPLITS, FAVORITES, JOURNAL, MemCirName, MemCirSets, MemSets, model, ModelParam, modt30times, OPTIONS, SetDescriptions, SPLITS, StrokeCategory, WkParam, WorkCategory, Workout, MTEVENTE, TestData, TestLine, TestSet, TestT30;");
	oConn.Execute ("DELETE TEAM.*, Athlete.Athlete FROM TEAM LEFT JOIN Athlete ON TEAM.Team = Athlete.Team1 WHERE (((Athlete.Athlete) Is Null));");
	
	oConn.Execute ("CREATE INDEX Birth ON Athlete (Birth);");
	oConn.Execute ("CREATE INDEX Sex ON Athlete (Sex);");
	oConn.Execute ("CREATE INDEX Team1 ON Athlete (Team1);");
	
	//oConn.Execute ("CREATE INDEX ATHLETE ON RESULT (ATHLETE);");
	oConn.Execute ("CREATE INDEX MEET ON RESULT (MEET);");
	oConn.Execute ("CREATE INDEX ATHLETE ON RESULT (COURSE);");
	oConn.Execute ("CREATE INDEX EVENTS ON RESULT (STROKE,DISTANCE);");
	oConn.Execute ("CREATE INDEX RANK ON RESULT (RANK);");
	oConn.Execute ("CREATE INDEX SCORE ON RESULT (SCORE);");
	
	oConn.Execute ("SELECT RESULT.COURSE, Int(DateDiff('m',[Athlete]![Birth],Date())/12) AS Birth, Athlete.Sex, RESULT.STROKE, RESULT.DISTANCE INTO ResultInfo FROM RESULT INNER JOIN Athlete ON RESULT.ATHLETE = Athlete.Athlete WHERE (((RESULT.RANK)=1) AND ((RESULT.I_R)='I') AND ((Athlete.Group)='A')) GROUP BY RESULT.COURSE, Int(DateDiff('m',[Athlete]![Birth],Date())/12), Athlete.Sex, RESULT.STROKE, RESULT.DISTANCE ORDER BY Athlete.Sex, RESULT.STROKE, RESULT.DISTANCE;");
	
	oConn.Execute ("CREATE TABLE RECBREAKERS (RecDate INTEGER,[RecFile] TEXT(8),RecTime INTEGER, RecText Text(80), Stroke INTEGER, Distance INTEGER, Lo_Age INTEGER, Hi_Age INTEGER);");	
	oConn.Execute ("ALTER TABLE RECNAME ADD COLUMN RecBroken INTEGER;");
	rs = oConn.Execute("SELECT RecFile FROM RECNAME;");
	var rst;
	while(!rs.eof)
	{
	oConn.Execute ("DROP INDEX Hi_Age ON [" + rs.Fields("RecFile") + "];");
	rst = oConn.Execute("SELECT Count(Rec) AS RecCount FROM [" + rs.Fields("RecFile") + "] HAVING ((Year(Date())-Year([RecDate])=0));");
	oConn.Execute("UPDATE RECNAME SET RECNAME.RecBroken = " + rst.Fields("RecCount") + " WHERE (((RECNAME.RecFile)='" + rs.Fields("RecFile") + "'));");
	rst.Close();
	
	rst = oConn.Execute("SELECT DateDiff('d',[RecDate],Date()) as ago, RecText, RecTime, Distance, Stroke, Lo_Age, Hi_Age FROM [" + rs.Fields("RecFile") + "] WHERE DateDiff('w',[RecDate],Date()) <= 2 And DateDiff('w',[RecDate],Date()) >=0 And I_R ='I';");
	while(!rst.eof){
	oConn.Execute("INSERT INTO RECBREAKERS VALUES (" + rst.Fields("ago") + ",'" + rs.Fields("RecFile") + "'," + rst.Fields("RecTime") + ",'" + rst.Fields("RecText") + "'," + rst.Fields("Stroke") + "," + rst.Fields("Distance") + "," + rst.Fields("Lo_Age") + "," + rst.Fields("Hi_Age") + ");");
	rst.MoveNext();
	}
	rst.Close();
	rs.MoveNext();
	
	}
	rs.Close();
	
	rs = oConn.Execute ("UPDATE STDNAME SET STDNAME.D1 = Trim([D1]), STDNAME.D2 = Trim([D2]),STDNAME.D3 = Trim([D3]),STDNAME.D4 = Trim([D4]),STDNAME.D5 = Trim([D5]),STDNAME.D6 = Trim([D6]),STDNAME.D7 = Trim([D7]),STDNAME.D8 = Trim([D8]),STDNAME.D9 = Trim([D9]),STDNAME.D10 = Trim([D10]),STDNAME.D11 = Trim([D11]),STDNAME.D12 = Trim([D12]);");
	rs = oConn.Execute ("SELECT STDNAME.StdFile, Trim([D1]) AS Expr1, Trim([D2]) AS Expr2, Trim([D3]) AS Expr3, Trim([D4]) AS Expr4, Trim([D5]) AS Expr5, Trim([D6]) AS Expr6, Trim([D7]) AS Expr7, Trim([D8]) AS Expr8, Trim([D9]) AS Expr9, Trim([D10]) AS Expr10, Trim([D11]) AS Expr11, Trim([D12]) AS Expr12 FROM STDNAME;");
	while(!rs.eof){ 
	RCOLOMS = "";
	for(c=0;c<12;c++)
	{
	if(rs.Fields(c+1) == "")
	{
	if(RCOLOMS != "")
	{
	RCOLOMS +=", ";
	}
	RCOLOMS += "[F(" + c + ")], [F(" + (c+12) + ")], [F(" + (c+24) + ")], [S(" + c + ")], [S(" + (c+12) + ")], [S(" + (c+24) + ")]";
	}
	}
	oConn.Execute ("ALTER TABLE ["+ rs.Fields("StdFile") +"] DROP COLUMN " + RCOLOMS + ";");
	oConn.Execute ("DROP INDEX Hi_Age ON ["+ rs.Fields("StdFile") +"];");
	oConn.Execute ("CREATE INDEX Event ON ["+ rs.Fields("StdFile") +"] (Stroke,Distance);");
	oConn.Execute ("CREATE INDEX Lo_Age ON ["+ rs.Fields("StdFile") +"] (Lo_Age);");
	oConn.Execute ("CREATE INDEX Hi_Age ON ["+ rs.Fields("StdFile") +"] (Hi_Age);");
	oConn.Execute ("CREATE INDEX I_R ON ["+ rs.Fields("StdFile") +"] (I_R);");
	oConn.Execute ("CREATE INDEX Sex ON ["+ rs.Fields("StdFile") +"] (Sex);");
	rs.MoveNext();
	}
rs.Close();
oConn.close();


//oConn.Execute ("DROP TABLE STDNAME;")

	
fs = Server.CreateObject("Scripting.FileSystemObject");
if(fs.FileExists("C:/Inetpub/wwwroot/Swim Database/extration/Swim4.mdb"))
{
fs.DeleteFile("C:/Inetpub/wwwroot/Swim Database/extration/Swim4.mdb");
}	
Engine = Server.CreateObject("JRO.JetEngine");
Engine.CompactDatabase(OpenString, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/Inetpub/wwwroot/Swim Database/extration/Swim4.mdb;Jet OLEDB:Database Password=");	
//fs.DeleteFile("C:/Inetpub/wwwroot/Swim Database/extration/" + database + ".mdb");
Engine = null;
fs = null;
//Response.Redirect ("Email_notify.asp?DBName=" + Request.QueryString("DBName") + "&Version=" + Request.QueryString("Version"));

%>

<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="en-za">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE>Times - Province</TITLE>

<link rel="stylesheet" type="text/css" href="../styles.css">

<script src="../menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(1);
</script>


</head>
<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="581" id="table1">
	<tr>
		<td align="center"><b><font size="5" color="#000080">Optimize and Fix 
		Database</font></b><p>&nbsp;</td>
	</tr>
</table>
</BODY></HTML><%}%>