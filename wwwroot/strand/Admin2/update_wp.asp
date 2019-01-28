<%@ Language=JavaScript %>
<%
Response.CacheControl = "no-store";
Response.Expires = 0;
if (Request.QueryString("DBName").count > 0 & Request.QueryString("Version").count > 0 & Request.QueryString("TMV").count > 0 & Request.QueryString("Message").count > 0)
{
var s, myFSO, WriteStuff;
s="";
s +="<%@ Language=JavaScript %" + ">"
s +="<%"
s +="Response.CacheControl = 'no-store';"
s +="Response.Expires = 0;"
s +="%" + ">"
s +="<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
s +="<HTML><HEAD><TITLE>Welcome To Western Province Swimming</TITLE>"
s +="<META http-equiv=Content-Type content='text/html; charset=windows-1252'>"
s +="<META HTTP-EQUIV='Pragma' CONTENT='no-cache'>"
s +=""
s +="<link rel='stylesheet' type='text/css' href='../styles.css'>"
s +=""
s +="<script src='../menu.asp' language='javascript' type='Text/Javascript'>"
s +="</script>"
s +=""
s +="<script language='javascript' type='Text/Javascript'>"
s +="makemenu(1);"
s +="</script>"
s +=""
s +=""
s +="<META content='Microsoft FrontPage 6.0' name=GENERATOR></HEAD>"
s +="<BODY bottomMargin=0 vLink=#ffffff link=#ffffff bgColor=#006498 leftMargin=0"
s +="topMargin=0 rightMargin=0>"
s +="<TABLE id=table1 cellSpacing=0 cellPadding=0 width=640 border=0>"
s +="<TBODY>"
s +="<TR>"
s +="<TD><IMG height=74 src='images/Logo.gif' width=666 border=0></TD></TR>"
s +="<TR>"
s +="<TD>&nbsp;"
s +="<DIV align=center>"
s +="<TABLE id=table2 width='65%' border=0>"
s +="<TBODY>"
s +="<TR>"
s +="<TD width='100%' height=91>"
s +="<P align=center><FONT face='Times New Roman' size=3><B>Latest"
s +="Western Province Team Manager II Data Base</B></FONT>"
s +="<P align=center><B><FONT face=Arial color=#ff0000>WPA Version " + Request.QueryString("Version") + " Update as on "
s += Date() + "</FONT></B><P align=center><FONT face='Times New Roman' size=3><B><FONT"
s +="color=#ff0000>&nbsp;</FONT></B></FONT><FONT"
s +="color=#ff0000>&nbsp;</FONT><B><FONT face=Arial color=#ff0000><a href='http://www.strandaquatics.za.net/western_province/secure/login.asp'>Download Now!!!</a></FONT></B></P></TD></TR>"
s +="<TR>"
s +="<TD width='100%'>"
s +="<P align=center><b>Team and Meet Manager security update for those"
s +="how are running Windows XP SP2<br>&nbsp;<a href='http://www.hy-tekltd.com/Updates/Hy-TekWindowsXPSP2Fix.exe'>Download"
s +="Now!!</a></b></P>"
s +="<P align=center><b><FONT face=Helvetica>All Meet entries to be"
s +="done on this Data Base only....</FONT></b></P>"
s +="<P align=center><b><font face='Helvetica'>Problems regarding"
s +="WP database please contact<br></font></b><font face='Helvetica' size='4'>&nbsp;<a href='mailto:lundi@discoverymail.co.za?subject=WP Database problem'>Lundi Bronkhorst</a></font></P>"
s +="<P>&nbsp;</P></TD></TR></TBODY></TABLE></DIV>"
s +="<P align=center>&nbsp;</P>"
s +="<P>&nbsp;</P></TD></TR></TBODY></TABLE></BODY></HTML>"

myFSO = Server.CreateObject("Scripting.FileSystemObject");

WriteStuff = myFSO.OpenTextFile("c:/Domains/strandaquatics.com/wwwroot/western_province/default.asp",2,true);
WriteStuff.WriteLine(s);

WriteStuff.Close();
WriteStuff = null;


s="";
s+="<%@ Language=JavaScript %" + ">"
s +="<%"
s +="Response.CacheControl = 'no-store';"
s +="Response.Expires = 0;"
s +="%" + ">"
s +="<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
s +="<HTML><HEAD><TITLE>Welcome To Western Province Swimming</TITLE>"
s +="<META http-equiv=Content-Type content='text/html; charset=windows-1252'>"
s +=""
s +="<link rel='stylesheet' type='text/css' href='../../styles.css'>"
s +=""
s +="<script src='../../menu.asp' language='javascript' type='Text/Javascript'>"
s +="</script>"
s +=""
s +="<script language='javascript' type='Text/Javascript'>"
s +="makemenu(2);"
s +="</script>"
s +=""
s +=""
s +="<META content='Microsoft FrontPage 6.0' name=GENERATOR></HEAD>"
s +="<BODY bottomMargin=0 vLink=#ffffff link=#ffffff bgColor=#006498 leftMargin=0"
s +="topMargin=0 rightMargin=0>"
s +="<TABLE id=table1 cellSpacing=0 cellPadding=0 width=640 border=0>"
s +="<TBODY>"
s +="<TR>"
s +="<TD><IMG height=74 src='http://www.strandaquatics.za.net/western_province/images/Logo.gif' width=666 border=0></TD></TR>"
s +="<TR>"
s +="<TD>&nbsp;"
s +="<DIV align=center>"
s +="<TABLE id=table2 width='65%' border=0>"
s +="<TBODY>"
s +="<TR>"
s +="<TD width='100%' height=91>"
s +="<P align=center><FONT face='Times New Roman' size=4><B>"
s +="<FONT"
s +="color=#ff0000>&nbsp;</FONT></B></FONT><FONT"
s +="color=#ff0000 size='4'>&nbsp;</FONT><B><FONT face=Arial color=#ff0000><a href='http://www.strandaquatics.za.net/western_province/secure/Sw" + Request.QueryString("TMV") + "Bkup" + Request.QueryString("DBName") + "-" + Request.QueryString("Version") + ".zip'><font size='5'>Download Now!!!</font></a></FONT></B></P></TD></TR>"
s +="<TR>"
s +="<TD width='100%'>"
s +="<P>&nbsp;</P></TD></TR></TBODY></TABLE></DIV>"
s +="<P align=center>&nbsp;</P>"
s +="<P>&nbsp;</P></TD></TR></TBODY></TABLE></BODY></HTML>"

myFSO = Server.CreateObject("Scripting.FileSystemObject");

WriteStuff = myFSO.OpenTextFile("c:/Domains/strandaquatics.com/wwwroot/western_province/secure/login.asp",2,true);
WriteStuff.WriteLine(s);

WriteStuff.Close();
WriteStuff = null;


myFSO = null;


	%>
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="en-za">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE>Updated W.P Pages</TITLE>

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
		<td align="center"><font size="5" color="#000080"><b>Updated W.P Pages</b></font><p>&nbsp;</td>
	</tr>
</table>
</BODY></HTML>
<%
Response.Redirect ("extra_database.asp?DBName=" + Request.QueryString("DBName") + "&Version=" + Request.QueryString("Version") + "&TMV=" + Request.QueryString("TMV")  + "&Message=" + Server.HTMLEncode(Request.QueryString("Message")));
}%>