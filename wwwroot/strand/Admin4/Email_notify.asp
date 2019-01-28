<%@ LANGUAGE="JavaScript"%>
<%
			Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.Email.value = 'Processing';\n");Response.write("/--></script>\n");

var Mail ,s, UpDate,sent,max,maxre;
UpDate = Date();

	var oConn,eurl;	
	var oRs;		
	var filePath;	
	sent = 0;
	max = 0;
	maxre = 0;
	filePath = Server.MapPath("../mail.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
		
	oRs = oConn.Execute ("SELECT * From Mail_List WHERE Email='strandaqua@telkomsa.net';");
	if(!oRs.eof)
	{
	while(!oRs.eof &  maxre < 10)
	{
	while(!oRs.eof)
	{

JMail = Server.CreateObject("JMail.SMTPMail");
JMail.ContentType = "text/html";
JMail.ServerAddress = "smtp.strandaquatics.za.net"
JMail.Sender = "notification@strandaquatics.za.net";
JMail.Subject = "Swimming Results Notification";
//JMail.AddRecipientEx(oRs.Fields("EMail") , Server.HTMLEncode(oRs.Fields("Name")) + ", " + Server.HTMLEncode(oRs.Fields("Surname")));
JMail.AddRecipientEx("strandaqua@telkomsa.net" , "Strand Test");

JMail.ReplyTo = "secretary@strandaquatics.za.net";
s = "";
s +="<HTML>";
s +="<HEAD>";
s +="<meta http-equiv='Content-Language' content='en-us'>";
s +="";
s +="<STYLE>";
s +="BODY {";
s +="font-family: Arial;";
s +="font-size: 12pt;";
s +="color: 000000;";
s +="margin-left: 135 px;";
s +="margin-top: 135 px;";
s +="background-position: top left;";
s +="background-repeat: no-repeat;";
s +="}";
s +="</STYLE>";
s +="</HEAD>";
s +="<BODY style='margin-left: 35 px;margin-top: 35 px;'>";
s +="<p align='left'><strong><font color='#000990' size='3'>Dear&nbsp;" + oRs.Fields("Name") + "&nbsp;" + oRs.Fields("Surname") + "</font></strong></p>";
s +="<p align='left'><strong><font color='#000990' size='3'>New Results and Times are ";
s +="now Available for Viewing.</font></strong></p>";
s +="<p align='left'><strong><font size='3' color='#000990'>Database Version:&nbsp;" + Request.Form("Version");
s +=" as of&nbsp;" + UpDate + " &nbsp; </font></strong></p>";
s +="<p align='left'><strong><font color='#000990' size='3'> " + Server.HTMLEncode(Request.Form("Message")) + " </p>";
s +="<p align='left'><strong><font color='#000990' size='3'>Visit:";
s +="<a href='http://www.strandaquatics.za.net'>http://www.strandaquatics.za.net</a></font></strong></p>";

s +="<p>Lastest Meet Results<br></p>";

s +="<p align='left'><strong><font color='#000990' size='3'>";
s +="<a href='http://www.strandaquatics.za.net/unsubscribe.asp?mail=" + oRs.Fields("EMail") + "'>Click Here to ";
s +="unsubscribe</a></font></strong></p>";
s +="</BODY>";
s +="</HTML>";


JMail.Body = s;

//JMail.AppendBodyFromFile("http://www.strandaquatics.za.net/notification.asp?Email=" + oRs.Fields("EMail") + "&Surname=" + oRs.Fields("Surname") + "&Name=" + oRs.Fields("Name") + "&Version=" + Request.Form("Version") + "&UpDate=" + Date());

//JMail.AddHeader ("Originating-IP", Request.ServerVariables("REMOTE_ADDR"));
//JMail.AddAttachment( "../Header_Strand.jpg" );
max++;
for(t = 0;t<4;t++)
{
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.Email.value = 'Sending try " + (t+1) + " of 4';\n");Response.write("/--></script>\n");

try{
JMail.Execute ();
JMail = null;
t = 10;
sent++;
break;break;
}
catch(System)
{

}
maxre++;
}

oRs.MoveNext();
}


}

Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.Email.value = '" + sent + " of " + max +" Sent successfully';\n");Response.write("/--></script>\n");

}%>