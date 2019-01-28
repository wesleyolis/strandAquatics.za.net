<%@ LANGUAGE="JavaScript"%>
<%Response.Expires = 0 %>

<%
var Mail;



	
	var oConn;	
	var oRs;		
	var filePath;	
	filePath = Server.MapPath("../mail.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
		
	rs = oConn.Execute ("SELECT * From Mail_List;");

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<title>E-Mail Notification List</title>
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
		<p align="center"><b><font size="4" color="#000990">E-Mail Notifications List of subscribed Members<br></font></b><div align="center">
			<table border="0" cellpadding="0" cellspacing="0" width="455" id="table2" style="font-size: 10pt">
				<tr class="Instruct">
					<td width="130"><font color="#000990">Surname</font></td>
					<td width="130"><font color="#000990">Name</font></td>
					<td><font color="#000990">E-Mail</font></td>
				</tr>
				<%while(!rs.eof)
				{%>
				<tr><td><%=rs.Fields("Surname")%></td><td><%=rs.Fields("Name")%></td><td><%=rs.Fields("EMail")%></td>
				</tr>
				<%rs.MoveNext();
				}%>
			</table>
		</div>
		</td>
	</tr>
</table>
	

</body>

</html>
