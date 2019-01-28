<%@ Language=JavaScript %>
<%
var webdir = "c:/Domains/strandaquatics.com/wwwroot/";
if( Request.Form("DBName").count > 0 & Request.Form("Version").count > 0  & Request.Form("TMV").count > 0 )
{
Response.write("<script language='javascript'><!--/\n");
Response.write ("document.up.WDB.value = 'Processing';\n");
Response.write("/--></script>\n");
	
	var pass = new Array( );
	var conn,oConn, Engine, fs;	
	var rs;		
	var filePath, OpenString;	
	var database = Request.Form("DBName");
	filePath = webdir + "extration/" + database + ".mdb";
	oConn = Server.CreateObject("ADODB.Connection");
	conn = false;
	if(Request.Form("TMV") != "TM4")
	{
		Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Connecting to TM2 DB..';\n");Response.write("/--></script>\n");
		OpenString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath;
		oConn.Open(OpenString);
		Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Connected to TM2 DB';\n");Response.write("/--></script>\n");
		if(conn == false)
		{
			Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Error Conneting to TM2 DB, contact wesley';\n");Response.write("/--></script>\n");
		}

	
	}
	else
	{
	Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Conneting to TM4 DB';\n");Response.write("/--></script>\n");

		for(i=0;i<5 & ( conn == false);i++)
		{
			try
			{
				OpenString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Jet OLEDB:Database   + ";";
				Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Try pass " + (i+1) + " of 5..';\n");Response.write("/--></script>\n");
				oConn.Open(OpenString);
				//reset the password on the database to nothing to be compatible with rest of site.
				conn = true;
			}
			catch(System)
			{
			
			}
		}
		
		if(conn == false)
		{
			Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Error Conneting to TM4 DB, contact wesley';\n");Response.write("/--></script>\n");
		}
		else
		{
			Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Connected to TM4 DB, Optermizing..';\n");Response.write("/--></script>\n");
		}


	}

	if(conn != false)
	{

	
	
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
	oConn.Execute ("ALTER TABLE RESULT DROP COLUMN SPLIT, EX, ORIGIN, MISC;");
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
	
	/*oConn.Execute ("CREATE TABLE RECBREAKERS (RecDate INTEGER,[RecFile] TEXT(8),RecTime INTEGER, RecText Text(120), Stroke INTEGER, Distance INTEGER, Lo_Age INTEGER, Hi_Age INTEGER);");	
	oConn.Execute ("ALTER TABLE RECNAME ADD COLUMN RecBroken INTEGER;");
	try{
	rs = oConn.Execute("SELECT RecFile FROM RECNAME;");
	var rst;
	while(!rs.eof)
	{
	oConn.Execute ("DROP INDEX Hi_Age ON [" + rs.Fields("RecFile") + "];");
	rst = oConn.Execute("SELECT Count(Rec) AS RecCount FROM [" + rs.Fields("RecFile") + "] HAVING (( IIf(Month(Date())<5,DateDiff('m',[RecDate],DateSerial((Year(Date())),5,1)),DateDiff('m',[RecDate],DateSerial((Year(Date())+1),5,1))) <=12) And ( IIf(Month(Date())<5,DateDiff('m',[RecDate],DateSerial((Year(Date())),5,1)),DateDiff('m',[RecDate],DateSerial((Year(Date())+1),5,1))) >0));");
	oConn.Execute("UPDATE RECNAME SET RECNAME.RecBroken = " + rst.Fields("RecCount") + " WHERE (((RECNAME.RecFile)='" + rs.Fields("RecFile") + "'));");
	rst.Close();
	
	rst = oConn.Execute("SELECT  IIf(InStr(1,[RecText],'at',1)>0,'<b>' & Left([RecText],InStr(1,[RecText],'at',1)-1) & '</b>' & Mid([RecText],InStr(1,[RecText],'at',1),Len([RecText])),[RecText]) AS RecTextB, DateDiff('d',[RecDate],Date()) as ago, RecText, RecTime, Distance, Stroke, Lo_Age, Hi_Age FROM [" + rs.Fields("RecFile") + "] WHERE (DateDiff('w',[RecDate],Date()) <= 8) And (DateDiff('w',[RecDate],Date()) >=0) And I_R ='I';");
	var cnt = 0;
	while((!rst.eof) & (cnt < 80)){
	oConn.Execute("INSERT INTO RECBREAKERS VALUES (" + rst.Fields("ago") + ",'" + rs.Fields("RecFile") + "'," + rst.Fields("RecTime") + ",'" + rst.Fields("RecTextB") + "'," + rst.Fields("Stroke") + "," + rst.Fields("Distance") + "," + rst.Fields("Lo_Age") + "," + rst.Fields("Hi_Age") + ");");
	rst.MoveNext();
	cnt++;
	}
	
	rst.Close();
	rs.MoveNext();
	
	
	}	rs.Close();}
	catch(System)
	{
	
	}*/

	
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
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Compacting-Repair DB';\n");Response.write("/--></script>\n");

	
fs = Server.CreateObject("Scripting.FileSystemObject");
if(fs.FileExists(webdir + "Swim.mdb"))
{
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Deleting online DB, if NO online,contact wesley';\n");Response.write("/--></script>\n");
fs.DeleteFile(webdir + "Swim.mdb");
}	
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'If No Online';\n");Response.write("/--></script>\n");

Engine = Server.CreateObject("JRO.JetEngine");
Engine.CompactDatabase(OpenString, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + webdir + "Swim.mdb;Jet OLEDB:Database Password=");	

Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Online DB Updated, Delete Temp';\n");Response.write("/--></script>\n");

fs.DeleteFile(webdir + "extration/" + database + ".mdb");


Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WDB.value = 'Online DB Updated and Running';\n");Response.write("/--></script>\n");




Engine = null;
fs = null;
}
}%>