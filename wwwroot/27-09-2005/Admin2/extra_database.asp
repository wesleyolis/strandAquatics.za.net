<%@ Language=VBScript %>
<%
Response.CacheControl = "no-store"
Response.Expires = 0
If Request.QueryString("DBName").count > 0 And Request.QueryString("Version").count > 0 And Request.QueryString("TMV").count > 0 Then
Dim objZip

Set objZip = Server.CreateObject("XStandard.Zip")
objZip.UnPack "c:\Domains\strandaquatics.com\wwwroot\western_province\secure\Sw" & Request.QueryString("TMV") & "Bkup" & Request.QueryString("DBName") & "-" & Request.QueryString("Version") & ".zip", "c:\Domains\strandaquatics.com\wwwroot\extration\"

Set objZip = Nothing


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
		<td align="center"><b><font size="5" color="#000080">Extra and Optimize and Fix 
		Database</font></b><p>&nbsp;</td>
	</tr>
</table>
</BODY></HTML><%
Response.Redirect ( "Optimize.asp?DBName=" & Request.QueryString("DBName") & "&Version=" & Request.QueryString("Version") & "&TMV=" & Request.QueryString("TMV") & "&Message=" & Server.HTMLEncode(Request.QueryString("Message")))
End IF%>