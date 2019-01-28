<%@ Language=JavaScript %>
<%
s = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>";
s +="<HTML><HEAD>";
s +="<META http-equiv=Content-Type content='text/html; charset=iso-8859-1'>";
s +="<META content='MSHTML 6.00.2900.2523' name=GENERATOR>";
s +="<STYLE>BODY { BACKGROUND-POSITION: left top; MARGIN-TOP: 135px; FONT-SIZE: 12pt; MARGIN-LEFT: 135px; COLOR: #000000; BACKGROUND-REPEAT: no-repeat; FONT-FAMILY: Arial}</STYLE>";
s +="</HEAD>";
s +="<BODY style=\"MARGIN-TOP: 135px; MARGIN-LEFT: 135px\" bgColor=#ffffff  background=\"http://www.strandaquatics.za.net/Header_Strand.jpg\">";
s +="<P align=left><STRONG><FONT color=#000990 size=3>Dear ";
s +="Subscriber</FONT></STRONG></P>";
s +="<P align=left><STRONG><FONT color=#000990 size=3>This is your notification ";
s +="to&nbsp;let you know, <BR>that there are new results and times available for ";
s +="viewing.</FONT></STRONG></P><FONT color=#000990 size=3><b><u>Message:</u><br><br>" + Request.Form("msg") + "</b></font>";
s +="<P align=left><STRONG><FONT color=#000990 size=3>Visit: <A";
s +=" href='http://www.strandaquatics.za.net'>www.strandaquatics.za.net</A></FONT></STRONG></P>";
s +="<P align=left><STRONG><FONT color=#000990 size=3>";
s +="<a href='http://www.strandaquatics.za.net/unsubscribe.asp'>Click Here to unsubscribe</A>";
s +="</FONT></STRONG></P>";
s +="<P align=left><STRONG><FONT color=#000990";
s +="size=3></FONT></STRONG>&nbsp;</P>";
s +="<TABLE style=\"RIGHT: 10px; LEFT: 132px; POSITION: absolute; TOP: 111px\" cellSpacing=0 cellPadding=0 width=643 align=left border=0>";
s +="<TBODY><TR><TD width=9>&nbsp;</TD><TD width=365><B>Web Site: <A href='http://www.strandaquatics.za.net' target=_blank>www.strandaquatics.za.net</A></B></TD>";
s +="<TD width=130><B><A href='http://strandaquatics.za.net/contact.htm'>Contact Us</A></B></TD>";
s +="<TD><B><A href='http://strandaquatics.za.net/meets/index.htm'>Meet Results</A></B></TD></TR></TBODY></TABLE>";
s +="</BODY></HTML>";


Mail = Server.CreateObject("CDO.Message");

Mail.From = "Notify Strand Aquatics <notification@strandaquatics.za.net>";
Mail.To = "hetty.olis@sazi.co.za"
Mail.Subject = "New Swimming Results Notification";
//Mail.CreateMHTMLBody("http://www.strandaquatics.za.net/notification.htm");
Mail.TextBody = "hello";
Mail.HTMLBody = s;
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2;
Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.strandaquatics.za.net";
Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25;
Mail.Configuration.Fields.Update();
try
{
Mail.Send();
Mail = null;
Response.write("Successfully\n");

}
catch(Exception)
{
Response.write("Error\n");

}



%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>

</body>

</html>
