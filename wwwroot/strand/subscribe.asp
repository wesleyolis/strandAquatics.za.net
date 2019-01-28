<%@ LANGUAGE="JScript"%> <%Response.Expires = 0 %> <%
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
%> <%
var Fill, Name, Surname, Mail,valid = true;
Fill = Request.Form("Fill");
if(Fill == "true")
{
Name = Request.Form("Name");
Surname =  Request.Form("Surname");
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
	if(document.subscribe.Surname.value.length >= 4)
	{
		if(document.subscribe.Name.value.length >= 4){

			return true;
			}else{alert("Name must be of a minimum length 4");}
	}else
	{alert("Surname must be of a minimum length 4");}
}
else
{
alert("Please enter a valid E-mail Address");

}return false; }

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
<body bgcolor="White" topmargin="0" leftmargin="0">

<%if(Fill != "true")
{%>
<form name="subscribe" method="POST" action="subscribe.asp" onsubmit="return validate()">
	<input type="hidden" name="Fill" value="true">
	<table border="0" cellpadding="0" cellspacing="0" width="470" id="table3">
		<tr>
			<td align="center">
			<p class="heading"><b>E-mail Notification</b></p>
			By subscribing to this e-mail system. <br>
			You will receive e-mail notification when new results/times are available.<br>
			<br>
			<table border="0" width="311" id="table4" cellspacing="0" cellpadding="0" height="127">
				<tr>
					<td width="70"><b>Name</b></td>
					<td width="252">
					<input name="Name" size="30" style="font-weight: 700" maxlength="20"></td>
				</tr>
				<tr>
					<td><b>Surname</b></td>
					<td width="252">
					<input name="Surname" size="30" style="font-weight: 700" maxlength="20"></td>
				</tr>
				<tr>
					<td><b>E-Mail</b></td>
					<td width="252">
					<input name="Mail" size="30" style="font-weight: 700" maxlength="64" value="<%if(Request.QueryString("mail").count !=0){%><%=Request.QueryString("mail") %><%}%>"></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td width="252">
					<input type="submit" value="Subscribe" name="sub" style="font-weight: 700; float: right"></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
</form>
<p><%}
else
{
%> </p>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="470" id="table5">
		<tr>
			<td align="center">
			<p class="heading"><b>E-mail Notification</b></p>
			<table border="1" cellpadding="5" cellspacing="0" width="329" id="table6" height="117" bordercolor="#000000">
				<tr>
					<td align="center"><b><%

	var oConn;	
	var oRs;		
	var filePath;	
	filePath = Server.MapPath("mail.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
		
	oRs = oConn.Execute ("SELECT * From Mail_List  WHERE (Email = '" + Mail + "')");
	if(oRs.eof)
	{
	try{
	oConn.Execute ("insert into Mail_List (Surname, Name, Email) values ('" + Surname + "','" + Name + "','" + Mail + "')");

%> Thank <% = Name + " " + Surname %> for Subscribing to our mail notification system.
					<%
}catch(Exception)
{%><a href="subscribe.asp">Error please try again!</a><%}
}
else
{
%> The E-Mail address &#39;<% =Mail %>&#39; is all ready registered here!<br>
					<br>
					Do you want <a href="unsubscribe.asp?mail=<%=Mail %>">Unsubscribe</a>
					<%
}}
%> </b></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
</div>

</body>

</html>
