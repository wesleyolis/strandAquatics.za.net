<%@ LANGUAGE="JavaScript"%>
<%Response.Expires = 0 %>

<%
var Mail,address,s,body;
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


var Surname = Server.HTMLEncode(oRs.Fields("Surname"));
var Name = oRs.Fields("Name");

var porp = "&Email=true&Surname=" + Surname + "&Name=" + "&Version=" + Request.QueryString("Version");

Mail = Server.CreateObject("CDO.Message");

Mail.From = "Strand Aquatics <strandaqua@telkomsa.net>";
Mail.To = "wesleyo@telkomsa.net";
Mail.Subject = "New Swimming Results Notification";
//Mail.HTMLBody = "";
eurl = "http://www.strandaquatics.za.net/notification.asp?" + porp;

Mail.CreateMHTMLBody(eurl);
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2;
Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.strandaquatics.za.net";
Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25;
Mail.Configuration.Fields.Update();
Mail.Send();
Mail = null;
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