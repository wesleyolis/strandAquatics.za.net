<%@ LANGUAGE="JavaScript"%>
<%Response.Expires = 0 %>
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="en-us">

<STYLE>
BODY {
font-family: Arial;
font-size: 12pt;
color: 000000;
margin-left: 135 px;
margin-top: 135 px;
background-position: top left;
background-repeat: no-repeat;
}
</STYLE>
</HEAD>
<BODY style="margin-left: 135 px;margin-top: 135 px;" background="Header_Strand.jpg">







<table border="0" cellpadding="0" cellspacing="0" width="643" style="position: absolute; left: 132px; right: 10; top: 111px" align="left">
	<tr>
		<td width="9">&nbsp;</td>
		<td width="365"><b>Web Site:
		<a target="_blank" href="http://strandaquatics.za.net">
		www.strandaquatics.za.net</a></b></td>
		<td width="130"><b>
		<a href="http://strandaquatics.za.net/contact.htm">Contact Us</a></b></td>
		<td ><b><a href="http://strandaquatics.za.net/meets/index.htm">Meet 
		Results</a></b></td>
	</tr>
</table>
<p>&nbsp;</p>
<p align="left"><strong><font color="#000990" size="3">Dear&nbsp;<%=Request.QueryString("Name")%>&nbsp;<%=Request.QueryString("Surname")%></font></strong></p>
<p align="left"><strong><font color="#000990" size="3">New Results and Times are 
now Available for Viewing.</font></strong></p>
<p align="left"><strong><font size="3" color="#000990">Database Version:&nbsp;<%=Request.QueryString("Version")%> 
as of&nbsp;<%=Request.QueryString("UpDate")%> &nbsp; </font></strong></p>
<p align="left"><strong><font color="#000990" size="3">Visit:
<a href="http://www.strandaquatics.za.net">http://www.strandaquatics.za.net</a></font></strong></p>
<p align="left"><strong><font color="#000990" size="3">
<a href="http://www.strandaquatics.za.net/unsubscribe.asp?Email=<%=Request.QueryString("Email")%>">Click Here to 
unsubscribe</a></font></strong></p>
</BODY>
</HTML>