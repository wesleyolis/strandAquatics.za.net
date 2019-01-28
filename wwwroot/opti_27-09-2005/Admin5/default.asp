<%@ Language=JavaScript %>
<%
Response.CacheControl = "no-store"
Response.Expires = 0;
Response.Buffer = false;

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
<%if( Request.Form("DBName").count == 0 & Request.Form("Version").count == 0  & Request.Form("TMV").count == 0 )
{%>
<table border="0" cellpadding="0" cellspacing="0" width="581" id="table1">
	<tr>
		<td align="center" valign="top"><font size="5" color="#000080"><b>Admin 
		- Update Database</b></font><form name="updb" method="POST" action="default.asp">
			<table border="0" cellpadding="0" cellspacing="0" width="471" id="table2" class="THeader">
				<tr>
					<td height="25" width="471" colspan="3" align="left" valign="top">
					Select Team Manager Version</td>
				</tr>
				
				<tr>
					<td height="26" width="118" align="left" valign="top">
					<input disabled type="radio" value="tm" name="TMV"><font color="#C0C0C0">TM 2.0</font></td>
					<td height="26" width="160" align="left" valign="top">
					<input type="radio" name="TMV" value="TM4" checked>TM 4.0</td>
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
				
					<textarea rows="1" name="Message" cols="56"></textarea>
					One line only..</td>
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
<%}else{%>

<table border="0" width="641" id="table3" cellspacing="0" cellpadding="0" height="190">
	<tr>
		<td align="center" height="40"><font size="5" color="#000080"><b>Admin 
		- Update Database</b></font></td>
	</tr>
	<tr>
		<td align="center">
		<form name="up" style="color: #000990; font-weight: bold">
			<table border="0" cellpadding="0" cellspacing="0" width="615" id="table4" style="font-size: 12pt;font-weight: bold; color:#000080">
			<tr>
				<td width="249" height="39"><font size="4">Updating in Process</font></td>
				<td height="39">&nbsp;</td>
			</tr>
			<tr>
				<td width="249">Western Province</td>
				<td> 
				<input type="text" name="WP" size="45" style="border-style:solid; border-width:0; background-color: #4698D3; color:#000990; font-weight:bold"></td>
			</tr>
			<tr>
				<td>Web Database</td>
				<td> 
				<input type="text" name="WDB" size="45" style="border-style:solid; border-width:0; background-color: #4698D3; color:#000990; font-weight:bold"></td>
			</tr>
			<tr>
				<td>Sending out Email notifications</td>
				<td> 
				<input type="text" name="Email" size="45" style="border-style:solid; border-width:0; background-color: #4698D3; color:#000990; font-weight:bold"></td>
			</tr>
		</table>
		</form>
		
		</td>
	</tr>
</table>



<%


Server.Execute("update_wp.asp");
//?DBName=" + Request.Form("DBName") + "&Version=" + Request.Form("Version") + "&TMV=" + Request.Form("TMV"));

Server.Execute("extra_database.asp");
//?DBName=" + Request.Form("DBName") + "&Version=" + Request.Form("Version") + "&TMV=" + Request.Form("TMV"));

Server.Execute("Optimize.asp");
//?DBName=" + Request.Form("DBName") + "&Version=" + Request.Form("Version") + "&TMV=" + Request.Form("TMV"));

//Server.Execute("Email_notify.asp");
//?DBName=" + Request.Form("DBName") + "&Version=" + Request.Form("Version") + "&TMV=" + Request.Form("TMV"));



}

%>



</BODY></HTML>