<%@ LANGUAGE="JScript"%>
<%Response.Expires = 0 %>
<%
function validText(text)
{
invalidChars = ""

for(i=0;i<text.length;i++)
		{
		charvalue = text.charAt(i)
		if(charvalue >2)
			{
			return false;
			}
		}
		
	return true;

}

function validEmail(email)
{
	invalidChars = " //:,;"
	if(email == "")
	{
	return false;
	}


	atpos = email.lastindexOf("@",1);

	if(atpos == -1)
	{
	return false;
	}
	
	periodpos = email.indexOf(".",atpos)

	if(periodpos == -1)
	{
	return false;
	}


	if(email.indexOf("@",atpos + 1) > -1)
	{
	return false;
	}
	
	if((periodpos +3) >email.length)
	{
	return false;
	}
	
		for(i=0;i<invalidChars.length;i++)
		{
		if(email.indexOf(invalidChars.charAt(i),0)>-1)
			{
			return false;
			}
		}
		
	return true;
}
%>

<%
var Fill, Mail,valid = true;
Fill = Request.Form("Fill");
if(Fill == "true")
{
Mail = Request.Form("Mail");

}
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<title>Subscribe</title>
</head>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);

function validate()
{
if(validEmail(document.subscribe.Mail.value))
{
return true;
}
else
{
alert("Please enter a valid E-mail Address");
return false; 
}}

function validEmail(email)
{
	invalidChars = " //:,;"
	if(email == "")
	{
	return false;
	}

	atpos = email.indexOf("@",1);

	if(atpos == -1)
	{
	return false;
	}	
	periodpos = email.indexOf(".",atpos)

	if(periodpos == -1)
	{
	return false;
	}


	if(email.indexOf("@",atpos + 1) > -1)
	{
	return false;
	}
	
	if((periodpos +3) >email.length)
	{
	return false;
	}
	
		for(i=0;i<invalidChars.length;i++)
		{
		if(email.indexOf(invalidChars.charAt(i),0)>-1)
			{
			return false;
			}
		}

	return true;
}

//-->
</script>

<body bgcolor="White" topmargin="0" leftmargin="0" >
<%if(Fill != "true" & Request.QueryString("Email").count == 0)
{%>

<form name="subscribe" method="POST" action="unsubscribe.asp"  onsubmit="return validate()">
	<input type="hidden" name="Fill" value="true">
	<table border="0" cellpadding="0" cellspacing="0" width="470" id="table3">
		<tr>
			<td align="center"><span  class="heading"><b>E-mail Notification</b></span>
			<br><br>please enter you e-mail address to unsubscribe<br><br>
			<table border="0" width="311" id="table4" cellspacing="0" cellpadding="0" height="87">
		<tr>
			<td><b>E-Mail</b></td>
			<td width="252">
			<input name="Mail" size="30" style="font-weight: 700" maxlength="64" value="<%if(Request.QueryString("mail").count !=0){%><%=Request.QueryString("mail") %><%}%>"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td width="252">
			&nbsp;</td>
		</tr>
		<tr>
			<td height="29">&nbsp;</td>
			<td width="252" height="29">
			<input type="submit" value="Unsubscribe" name="sub" style="font-weight: 700; float: right"></td>
		</tr>
	</table>
			</td>
		</tr>
	</table>
</form>

<%}
else
{
%>
<div align="center">
<div align="left">
	<table border="0" cellpadding="0" cellspacing="0" width="470" id="table5">
		<tr>
			<td align="center"><span  class="heading"><b>E-mail Notification</b></span>
			<br><br>
<table border="1" cellpadding="5" cellspacing="0" width="329" id="table6" height="117" bordercolor="#000000">
	<tr>
		<td align="center"><b>
<%
if(Request.QueryString("Email").count !=0)
{
Mail = Request.QueryString("Email");
}

	var oConn;	
	var oRs;		
	var filePath;	
	filePath = Server.MapPath("mail.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
		
	oRs = oConn.Execute ("SELECT * From Mail_List  WHERE (Email = '" + Mail + "')");
	if(!oRs.eof)
	{
	try{
	%>
		Sorry to see you go <%=oRs.Fields("Name")%>, <%=oRs.Fields("Surname")%> .
	<%
	oConn.Execute ("DELETE FROM Mail_List WHERE (Email = '" + Mail + "')");

} catch(Exception)
{%><a href="unsubscribe.asp">Error please try again!</a><%}
}
else
{
%>
The E-Mail address '<% =Mail %>' is not registered here!<br><br>Do you 
		want <a href="subscribe.asp?mail=<% =Mail %>">Subscribe</a>
<%
}}
%>
</td>
	</tr>
</table>
			</td>
		</tr>
	</table>
</div>
</div>
</body>

</html>