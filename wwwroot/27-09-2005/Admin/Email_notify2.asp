<%@ LANGUAGE="JavaScript"%>

<%
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.Email.value = 'Processing';\n");Response.write("/--></script>\n");


var Mail,address,s,body;
body = "Hallo"
s = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>";
s +="<HTML><HEAD>";
s +="<META http-equiv=Content-Type content='text/html; charset=iso-8859-1'>";
s +="<META content='MSHTML 6.00.2900.2523' name=GENERATOR>";
s +="<STYLE></STYLE>";
s +="</HEAD>";
s +="<BODY bgColor=#ffffff>";
s +="<P align=left><STRONG><FONT color=#000990 size=4>Dear ";
s +="Subscriber</FONT></STRONG></P>";
s +="<P align=left><STRONG><FONT color=#000990 size=4>This is your notification ";
s +="to&nbsp;let you know, <BR>that there are new results and times available for ";
s +="viewing.</FONT></STRONG></P>";
s +="<P align=left><STRONG><FONT color=#000990 size=4>Visit: <A";
s +="href='http://www.strandaquatics.za.net'>www.strandaquatics.za.net</A></FONT></STRONG></P>";
s +="<P align=left><STRONG><FONT color=#000990 size=4><a href='www.strandaquatics.za.net'>bob</a>";
s +="<a href='http://www.strandaquatics.za.net/unsubscribe.asp'>Click Here to unsubscribe</A>";
s +="</FONT></STRONG></P>";
s +="<P align=left><STRONG><FONT color=#000990";
s +="size=4></FONT></STRONG>&nbsp;</P></BODY></HTML>";
body =s;


	var oConn;	
	var oRs;		
	var filePath;	
	filePath = Server.MapPath("../mail.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
		
	oRs = oConn.Execute ("SELECT EMail From Mail_List;");
	if(!oRs.eof)
	{
	while(!oRs.eof)
	{
	Mail = Server.CreateObject("CDO.Message");

	for(i = 0,address="";i<300;i++,oRs.MoveNext())
	{
	if(oRs.eof)
	{
	break;
	}
	address += oRs.Fields("EMail") + ";";
	}

	


Mail.From = "Notify Strand Aquatics <notification@strandaquatics.za.net>";
Mail.To = address
Mail.Subject = "New Swimming Results Notification";
//Mail.HTMLBody = "";
Mail.CreateMHTMLBody("http://www.strandaquatics.za.net/notification.htm");
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2;
Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.strandaquatics.za.net";
Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25;
Mail.Configuration.Fields.Update();
try
{
Mail.Send();
Mail = null;
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.Email.value = 'Sent successfully';\n");Response.write("/--></script>\n");

}
catch(Exception)
{
Response.write("<script language='javascript'><!--/\n");Response.write ("document.up.Email.value = '<a href='Email_notify.asp'>Failed Click here to attempt again</a>';\n");Response.write("/--></script>\n");

}

}



}%>