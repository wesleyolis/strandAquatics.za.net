<%@ LANGUAGE="JavaScript"%>
<%Response.Expires = 0 %>

<%






var Mail ,s, UpDate;
UpDate = Date();

	var oConn,eurl;	
	var oRs;		
	var filePath;	
	filePath = Server.MapPath("../mail.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
		
	oRs = oConn.Execute ("SELECT * From Mail_List WHERE Email='strandaqua@telkomsa.net';");
	if(!oRs.eof)
	{
	
	while(!oRs.eof)
	{

JMail = Server.CreateObject("JMail.SMTPMail");
JMail.ContentType = "text/html";
JMail.ServerAddress = "smtp.strandaquatics.za.net"
JMail.Sender = "secretary@strandaquatics.za.net";
JMail.Subject = "Swimming Results Notification";
JMail.AddRecipientEx(oRs.Fields("EMail") , Server.HTMLEncode(oRs.Fields("Name")) + ", " + Server.HTMLEncode(oRs.Fields("Surname")));
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
s +="<p align='left'><strong><font size='3' color='#000990'>Database Version:&nbsp;" + Request.QueryString("Version");
s +=" as of&nbsp;" + UpDate + " &nbsp; </font></strong></p>";
s +="<p align='left'><strong><font color='#000990' size='3'> " + Request.QueryString("Message") + " </p>";
s +="<p align='left'><strong><font color='#000990' size='3'>Visit:";
s +="<a href='http://www.strandaquatics.za.net'>http://www.strandaquatics.za.net</a></font></strong></p>";
s +="<p align='left'><strong><font color='#000990' size='3'>";
s +="<a href='http://www.strandaquatics.za.net/unsubscribe.asp?mail=" + oRs.Fields("EMail") + "'>Click Here to ";
s +="unsubscribe</a></font></strong></p>";
s +="</BODY>";
s +="</HTML>";


JMail.Body = s;

//JMail.AppendBodyFromFile("http://www.strandaquatics.za.net/notification.asp?Email=" + oRs.Fields("EMail") + "&Surname=" + oRs.Fields("Surname") + "&Name=" + oRs.Fields("Name") + "&Version=" + Request.QueryString("Version") + "&UpDate=" + Date());

//JMail.AddHeader ("Originating-IP", Request.ServerVariables("REMOTE_ADDR"));
//JMail.AddAttachment( "../Header_Strand.jpg" );

JMail.Execute ();
JMail = null;

oRs.MoveNext();
}


%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<title>E-Mail Notification</title>
</head>

<link rel="stylesheet" type="text/css" href="../styles.css">

<script src="../menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(1);
</script>

<body bgcolor="White" topmargin="0" leftmargin="0" >
	

<table border="0" cellpadding="0" cellspacing="0" width="470" id="table1">
	<tr>
		<td>
		<p align="center">&nbsp;</p>
		<p align="center"><b><font size="4" color="#000990">E-Mail Notifications 
		have been sent successfully</font></b></td>
	</tr>
</table>
	

</body>

</html>
<%}%>