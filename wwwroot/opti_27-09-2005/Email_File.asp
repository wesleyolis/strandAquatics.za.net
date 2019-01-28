<%@ LANGUAGE="JavaScript"%>
<%Response.Expires = 0 %>

<%
if(Request.QueryString("File").Count() == 0 || Request.QueryString("File") == "")
{%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>E-Mail File Request</title>
</head>

<body>
<form method="GET" action="http://strandaquatics.za.net/Email_File.asp">
	<table border="0" width="408" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td>File address:
	<input type="text" name="File" size="44">&nbsp;<br>E-Mail Address:
			<input type="text" name="EMail" size="41"><br>
			<input type="submit" value="Request File" name="Submit" style="float: right"><input type="reset" value="Reset" name="Reset" style="float: right"></td>
		</tr>
	</table>
	</p>
</form>
</body>

</html>

<%
}else{

var JMail;
	

JMail = Server.CreateObject("JMail.SMTPMail");

JMail.ServerAddress = "smtp.strandaquatics.za.net"


JMail.Sender = "FileMail@strandaquatics.za.net";
JMail.Subject = "File Mail";
JMail.AddRecipient (Request.QueryString("EMail"));

JMail.Body = "Attached find the file you requested File, please remeber to rename the extention URL: " + Request.QueryString("File");

JMail.AddHeader ("Originating-IP", Request.ServerVariables("REMOTE_ADDR"));

JMail.AddURLAttachment( Request.QueryString("File") , "Requested Attachment.efl" );

JMail.Execute ();

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>E-Mail File Request Conformation</title>
</head>

<body>
Mail Sent to successfully..<br>&nbsp;Should receive it soon<br>Remember to rename file 
immediately after save it .
</body>

</html>

<%}%>