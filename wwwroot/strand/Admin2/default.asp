<%@ Language=VBScript %>
<%
Response.CacheControl = "no-store"
Response.Expires = 0
	%>
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="en-za">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE>Administrator</TITLE>

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
		<td align="center" valign="top"><font size="5" color="#000080"><b>Admin 
		- Update Database</b></font><form method="GET" action="update_wp.asp">
			<table border="0" cellpadding="0" cellspacing="0" width="471" id="table2" class="THeader">
				<tr>
					<td height="25" width="471" colspan="3" align="left" valign="top">
					Select Team Manager Version</td>
				</tr>
				
				<tr>
					<td height="26" width="118" align="left" valign="top">
					<input type="radio" value="tm" checked name="TMV">TM 2.0</td>
					<td height="26" width="160" align="left" valign="top">
					<input type="radio" name="TMV" value="TM4">TM 4.0</td>
					<td height="26" width="188" align="left" valign="top">&nbsp;</td>
				</tr>
				
				<tr>
					<td height="29" width="283" colspan="2">Database Name&nbsp; <input type="text" name="DBName" size="20"> 
					</td>
					<td height="29" width="188" valign="top">Version&nbsp;
			<input type="text" name="Version" size="4"></td>
				</tr>
				
				<tr>
					<td height="34" colspan="3" valign="bottom"><u>Email 
					Notification Message </u><span style="font-weight: 400">
					<font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="Email_sub.asp">List of subscribe members</a></font></span></td>
				</tr>
				
				<tr>
					<td height="173" align="left" valign="top" colspan="3">
					<textarea rows="10" name="Message" cols="56"></textarea></td>
				</tr>
				<tr>
					<td colspan="3" align="right">
			<p>
			<input type="submit" value="Proceed" name="B1" style="font-weight: 700">&nbsp;&nbsp;
			<input type="reset" value="Reset" name="B2" style="font-weight: 700"></p>
					</td>
				</tr>
			</table>
		</form>
		<p>&nbsp;</td>
	</tr>
</table>
</BODY></HTML>