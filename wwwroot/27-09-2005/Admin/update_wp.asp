<%@ Language=JavaScript %>
<%
var webdir = "c:/Domains/strandaquatics.com/wwwroot/";
if (Request.Form("DBName").count > 0 & Request.Form("Version").count > 0 & Request.Form("TMV").count > 0)
{
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WP.value = 'Processing..';\n");Response.write("/--></script>\n");
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
s +="Western Province Team Manager " + Request.Form("TMV") + " Data Base</B></FONT>"
s +="<P align=center><B><FONT face=Arial color=#ff0000>WPA Version " + Request.Form("Version") + " Update as on "
s += Date() + "</FONT></B><P align=center><FONT face='Times New Roman' size=3><B><FONT"
s +="color=#ff0000>&nbsp;</FONT></B></FONT><FONT"
s +="color=#ff0000>&nbsp;</FONT><B><FONT face=Arial color=#ff0000><a href='http://www.strandaquatics.za.net/western_province/secure/login.asp'>Download Now!!!</a></FONT></B></P></TD></TR>"
s +="<TR><TD width='100%' align='center' height='49'><a href='http://www.strandaquatics.za.net/western_province/secure/login.asp'>Download WP TM4 Registration Database</a></TD></TR><TR>"
s +="<TR><TD align='center' height='49'><a href='TMHelp.htm'>Click Here for the WP TM4, doing your.<br>Registration and Meet Entries ect.</a></TD></TR><TR>"
s +="<TR>"
s +="<TD width='100%'>"
s +="<P align=center><b><FONT face=Helvetica>All Meet entries to be"
s +=" done on this Data Base only....</FONT></b></P>"
s +="<P align=center><b><font face='Helvetica'>Problems regarding"
s +=" WP database please contact<br></font></b><font face='Helvetica' size='4'>&nbsp;<a href='mailto:brianer@icon.co.za?subject=WP Database problem'>Sandra Reynolds </a></font></P>"
s +="<P>&nbsp;</P></TD></TR></TBODY></TABLE></DIV>"
s +="<P align=center>&nbsp;</P>"
s +="<P>&nbsp;</P></TD></TR></TBODY></TABLE></BODY></HTML>"

myFSO = Server.CreateObject("Scripting.FileSystemObject");
//WriteStuff = myFSO.OpenTextFile(webdir + "extration/default.asp",2,true);
WriteStuff = myFSO.OpenTextFile(webdir + "western_province/default.asp",2,true);
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
s +="color=#ff0000 size='4'>&nbsp;</FONT><B><FONT face=Arial color=#ff0000><a href='http://www.strandaquatics.za.net/western_province/secure/Sw" + Request.Form("TMV") + "Bkup" + Request.Form("DBName") + "-" + Request.Form("Version") + ".zip'><font size='5'>Download Now!!!</font></a></FONT></B></P></TD></TR>"
s +="<TR><TD align='center'><a href='http://www.strandaquatics.za.net/western_province/secure/SwTM4BkupWPA2005-2006Registr-01.zip'>WP - Registration Database</a></TD></TR>"
s +="</TBODY></TABLE></DIV>"
s +="<P align=center>&nbsp;</P>"
s +="<P>&nbsp;</P></TD></TR></TBODY></TABLE></BODY></HTML>"

myFSO = Server.CreateObject("Scripting.FileSystemObject");
//WriteStuff = myFSO.OpenTextFile(webdir + "extration/login.asp",2,true);

WriteStuff = myFSO.OpenTextFile(webdir + "western_province/secure/login.asp",2,true);
WriteStuff.WriteLine(s);

WriteStuff.Close();
WriteStuff = null;


myFSO = null;

Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WP.value = 'Updated';\n");Response.write("/--></script>\n");

}else
{
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.WP.value = 'Error Missing Parameters'\n;");Response.write("/--></script>\n");
}%>