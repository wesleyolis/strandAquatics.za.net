<%@ Language=JavaScript %>
<%
Response.CacheControl = "no-store";
Response.Expires = 0;


	var oConn, Engine, fs;	
	var rs;		
	var filePath, OpenString;	
	filePath = "c:/Domains/strandaquatics.com/wwwroot/Swim.mdb";

	OpenString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath;

	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open(OpenString);
	
	oConn.Execute("DELETE * FROM RECBREAKERS;");
	
	rs = oConn.Execute("SELECT RecFile FROM RECNAME;");
	var rst;
	while(!rs.eof)
	{
	
		rst = oConn.Execute("SELECT  IIf(InStr(1,[RecText],'at',1)>0,'<b>' & Left([RecText],InStr(1,[RecText],'at',1)-1) & '</b>' & Mid([RecText],InStr(1,[RecText],'at',1),Len([RecText])),[RecText]) AS RecTextB, DateDiff('d',[RecDate],Date()) as ago, RecText, RecTime, Distance, Stroke, Lo_Age, Hi_Age FROM [" + rs.Fields("RecFile") + "] WHERE (DateDiff('w',[RecDate],Date()) <= 8) And (DateDiff('w',[RecDate],Date()) >=0) And I_R ='I';");
		while(!rst.eof)
		{
		oConn.Execute("INSERT INTO RECBREAKERS VALUES (" + rst.Fields("ago") + ",'" + rs.Fields("RecFile") + "'," + rst.Fields("RecTime") + ",'" + rst.Fields("RecText") + "'," + rst.Fields("Stroke") + "," + rst.Fields("Distance") + "," + rst.Fields("Lo_Age") + "," + rst.Fields("Hi_Age") + ");");
		rst.MoveNext();
		}
		rs.MoveNext();
	}
	rst.Close();
	rs.Close();
oConn.close();

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
</BODY></HTML><%%>