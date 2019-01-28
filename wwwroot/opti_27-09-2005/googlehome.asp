<%@ Language=JavaScript %>
<%
Response.CacheControl = "Public";
Response.Expires = 1440;

function CTime (Score)
{
if(Score !=0)
{
if(Score<6000)
{t = (Score/100);
if(t <10.0)
{t ="0"+t;}
if(t.toString().length == 2)
{t = t + ".00";}else{
if(t.toString().length<5)
{t=t+"0";}}
}
else
{
s = (Score % 100);
if(s==0){s=".00"}else{
if(s <10){s = "0" + s;}}
if(s.toString().length<2){s = s + "0";}
Score = ((Score -s) / 100);
m = (Score % 60);
if(m < 10)
{m = "0" + m;}
Score = ((Score - m)/60);
h = (Score % 60);
if(h.toString().length<2){h = "0" + h;}
t = h + ":" + m + "." + s;}}
else{t = "D/Q";}
return t;}

function Age(Lo,Hi)
{if(Lo <  Hi)
{
if(Lo > 0 & Hi < 99)
{
age = Lo + " - " + Hi + "yrs";
}else
{
if(Hi == 99 & Lo >0)
{
age = Lo + "&nbsp;&amp;&nbsp;Over";
}
else
{
if(Lo == 0 & Hi < 99)
{
age = Hi + "&nbsp;&amp;&nbsp;Under";
}
else
{
if(Lo == 0 & Hi == 99)
{
age = "Open";
}}}}}else{age = Hi + "yrs";}

return age;
}

var strokes = new Array("","Freestyle","Back","Breast","Butterfly","Medley");
	var oConn;	
	var rst;		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	//rst = oConn.Execute("SELECT * FROM RECBREAKERS ORDER BY RecDate");
	
	var rst = oConn.Execute("SELECT RECNAME.RecFile, RECORDS.Record, * FROM RECNAME INNER JOIN RECORDS ON RECNAME.RECORD = RECORDS.Record WHERE (((DateDiff('w',[RecDate],Date()))<=4 And (DateDiff('w',[RecDate],Date()))>=0) AND ((RECORDS.I_R)='I')) ORDER BY RECORDS.RecDate;");
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">

<title>Strand Aquatics Swimming Club</title>
<link rel="stylesheet" type="text/css" href="styles.css">
<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>
<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>
</head>

<body topmargin="0" leftmargin="0">
<%if(!rst.eof){var pos = 0;%>

<table border="4" cellpadding="0" width="274" style="border-collapse: collapse; font-size:8pt" bordercolor="#000080" bordercolorlight="#0060BF" bgcolor="#6699FF" height="450" onmouseover="over = true;" onmouseout="over=false;">
<tr><td align="center" height="50"><b><font color="#000080" size="4">Congratulations to the following<br>&nbsp; Record Breakers</font></b></td></tr>
<tr><td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="266" id="table8">

<%while(!rst.eof){%>
<tr onmouseout="this.style.backgroundColor='';"  onmouseover="this.style.backgroundColor='#77A9FF';" onClick="self.location.replace('records.asp?recordFile=<%=rst.Fields("RECORDS.Record")%>&RecName=<%=rst.Fields("RecFile")%>&Lo=<%=rst.Fields("Lo_Age")%>&Hi=<%=rst.Fields("Hi_Age")%>');" id="rec<%=pos%>" <%if(pos > 7){%>style="display:none;"<%}%>><td class="r" height="37" valign="top">
<table border="0" cellpadding="2"  width="266" cellspacing="0" style="font-size: 8pt;">
<tr><td  valign="top" align="left" style="font-size: 10pt"><%=Age(rst.Fields("Lo_Age"),rst.Fields("Hi_Age"))%><br><%=CTime(rst.Fields("RecTime"))%></td>
<td  valign="top" width="203"><%=rst.Fields("Distance")%>m&nbsp;<%=strokes[rst.Fields("Stroke")]%>&nbsp;<%=rst.Fields("RecFile")%>&nbsp;Record,<br><%=rst.Fields("RecText")%>&nbsp;
</td></tr></table>
</td></tr>
<%rst.MoveNext();pos++;}
rst.Close();
%>

</table>
</td></tr>

</table>
&nbsp;<script language="javascript" defer>
<!--
var rec = 8,max = <%=pos%>, c, last = 0,f=0,over=false;
setTimeout("change()",4000);
function change()
{
if(over==false){
c = 0;
slideof();
//chck();
}
else
{
setTimeout('change()',2000);
}
}

function slideof()
{
if(c < 8)
{
	if( last < max)
	{
	document.getElementById("rec" + last).style.display = 'none';
	last++;
	}
	c++;
	//setTimeout("slideof()",80);
	slideof();
}
else
{
	if( last >= max)
	{last=0;}
	c=0;
	setTimeout("slideon()",100);
}
}

function slideon()
{
if(c < 8)
{
		if(rec < max)
	{
	document.getElementById("rec" + rec).style.display = 'block';
	rec++;
	}
	c++;
	setTimeout("slideon()",50);
	//slideon();
}
else
{
	if( rec >= max)
	{rec=0;}
	c=0;
	setTimeout("change()",4000);
}
}


function chck()
{
if(c < 8)
{
	if( last < max)
	{
	document.getElementById("rec" + last).style.display = 'none';
	last++;
	}
	
	if(rec < max)
	{
	document.getElementById("rec" + rec).style.display = 'block';
	rec++;
	}
	c++;
	setTimeout("chck()",300);
}
else
{
	if( last >= max)
	{last=0;}
	if( rec >= max)
	{rec=0;}
	setTimeout("change()",3000);
}
}
//-->
</script><%}%>
</body>
</html>