<%@ Language=JavaScript %>
<%
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

Response.Expires = -1;
var strokes = new Array("","Freestyle","Back","Breast","Butterfly","Medley");
	var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
%>
<%
	if(Request.QueryString("recordFile").count == 0)
	{
	rs = oConn.Execute("SELECT RECNAME.RECORD, RECNAME.RecFile, RECNAME.Year, RECNAME.Descript, RECNAME.Flag, RECNAME.Course FROM RECNAME ORDER BY RECNAME.Year DESC, RECNAME.RecFile;");
	var rst = oConn.Execute("SELECT RECNAME.RecFile, RECORDS.Record, * FROM RECNAME INNER JOIN RECORDS ON RECNAME.RECORD = RECORDS.Record WHERE (((DateDiff('w',[RecDate],Date()))<=4 And (DateDiff('w',[RecDate],Date()))>=0) AND ((RECORDS.I_R)='I')) ORDER BY RECORDS.RecDate;");
	
	var pos = 0;
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Records</title>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>

<style>
<!--
.r           { border-bottom: 1px solid #0066CC;cursor:hand }
-->
</style>


</head>
<body  topmargin="0" leftmargin="0" style="text-align: center">
<table border="0" cellpadding="0" cellspacing="0" width="657" id="table1" style="font-weight: bold">
<tr><td valign="top"><div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="803"><tr>
<td align="center" valign="top">
<table border="0" cellpadding="0" cellspacing="0" id="table1" style="font-weight: bold">
<tr><td colspan="4" align="center" height="53" valign="bottom"><b>
<font size="6" color="#000080" face="Arial Rounded MT Bold">Records Archives</font></b></td></tr>
<tr style="font-size:8pt;" valign="bottom"><td width="114" height="33">&nbsp;</td>
	<td width="188" height="33">&nbsp;</td><td width="58" height="33">&nbsp;</td>
<td  align="center" width="72" height="33" valign="top"><font color="#000080" size="2"><!--Broken<br>&nbsp;This 
Season--></font></td></tr>
<%while(!rs.eof){%>
<tr><td><a href="records.asp?recordFile=<%=rs.Fields("RECORD")%>&RecName=<%=rs.Fields("RecFile")%>"><%=rs.Fields("RecFile")%></a></td><td><%=rs.Fields("Descript")%></td><td><%=rs.Fields("Year")%></td><td align="center"></td></tr><%rs.MoveNext()}%></table>
</td>
<%if(!rst.eof){var pos = 0;%>
<td width="322" valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="280"height="51"><tr><td>&nbsp;</td></tr></table>

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

</td></tr></table>

<%}else{
var recFile = Request.QueryString("recordFile");
if(Request.QueryString("Lo").count == 0 || Request.QueryString("Hi").count == 0 || Request.QueryString("Hi") > 99 || Request.QueryString("Li") > 99){
//rs = oConn.Execute("SELECT [" + recFile + "].Lo_Age, [" + recFile + "].Hi_Age, Count([" + recFile + "].Rec) AS Broken FROM [" + recFile + "] WHERE ((Year([" + recFile + "]![RecDate])-Year(Date())=0)) GROUP BY [" + recFile + "].Lo_Age, [" + recFile + "].Hi_Age ORDER BY [" + recFile + "].Hi_Age;");
//rs = oConn.Execute("SELECT RECORDS.Lo_Age, RECORDS.Hi_Age, Count(RECORDS.Lo_Age) AS CountOfLo_Age FROM RECORDS WHERE (((((Month(Date()))<5) And ((DateDiff('m',[RecDate],DateSerial((Year(Date())),5,1)))<=12 And (DateDiff('m',[RecDate],DateSerial((Year(Date())),5,1)))>0) Or (((Month(Date()))>=5) And ((DateDiff('m',[RecDate],DateSerial((Year(Date())+1),5,1)))<=12))))) GROUP BY RECORDS.Lo_Age, RECORDS.Hi_Age, RECORDS.Record HAVING (((RECORDS.Record)=" +recFile + "));");
rs = oConn.Execute("SELECT RECORDS.Hi_Age, RECORDS.Lo_Age, '' As Broken FROM RECORDS GROUP BY RECORDS.Hi_Age, RECORDS.Lo_Age, RECORDS.Record HAVING (((RECORDS.Record)=" + recFile + ")) ORDER BY RECORDS.Hi_Age, RECORDS.Lo_Age;");

var lower, higher;
if(!rs.eof)
{
lower = rs.Fields("Lo_Age") + "";
higher = rs.Fields("Hi_Age")+ "" ;
first = "<tr><td></td><td><a href='records.asp?recordFile=" + recFile + "&RecName=" + Request.QueryString("RecName") + "&Lo=" + lower + "&Hi=" + higher + "'>" + Age(lower,higher) + "</td><td align='center'>" + rs.Fields("Broken") + "</td></tr>";
rs.MoveNext();
if(!rs.eof){
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=Request.QueryString("RecName")%>, All Records</title>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body  topmargin="0" leftmargin="0" style="text-align: center">
<table border="0" cellpadding="0" cellspacing="0" width="657" id="table1" style="font-weight: bold">
<tr><td valign="top"><div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="276" id="table2" style="font-weight: bold">
<tr><td align="center" height="63" colspan="2">
<font color="#000080" size="6" face="Arial Rounded MT Bold"><%=Request.QueryString("RecName")%></font><br><br><b>Select an Age Group Please</b></td>
	<td valign="bottom" align="center"><font color="#000080"><!--Broken--></font></td>
</tr><tr height="10"><td width="65" height="8"></td><td width="141" height="8"></td>
	<td width="70"></td></tr>
<%
Response.write(first);
while(!rs.eof){%><tr><td></td><td><a href="records.asp?recordFile=<%=recFile%>&RecName=<%=Request.QueryString("RecName")%>&Lo=<%=rs.Fields("Lo_Age")%>&Hi=<%=rs.Fields("Hi_Age")%>"><%=Age(rs.Fields("Lo_Age"), rs.Fields("Hi_Age"))%></a></td><td align="center"><%=rs.Fields("Broken")%></td></tr>
<%rs.MoveNext();
}

}
else
{
Response.Redirect("records.asp?recordFile=" + recFile + "&RecName=" + Request.QueryString("RecName") + "&Lo=" + lower + "&Hi=" + higher);
}
}else
{
%>No records found Sorry<%
}%></table><%}else{
var gender;
rs = oConn.Execute("SELECT Sex, Stroke, Distance, RecText, RecDate, IIf(Month(Date())<5,DateDiff('m',[RecDate],DateSerial((Year(Date())),5,1)),DateDiff('m',[RecDate],DateSerial((Year(Date())+1),5,1))) AS ago, RecTime, RecTeam, Ex FROM [RECORDS] WHERE (Record=" + recFile + ") And (((Lo_Age)="+Request.QueryString("Lo")+") AND ((Hi_Age)="+Request.QueryString("Hi")+")) ORDER BY Sex, Stroke, Distance;");
%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=Request.QueryString("RecName")%>,<%=Age(Request.QueryString("Lo"),Request.QueryString("Hi"))%> - Records</title>

<link rel="stylesheet" type="text/css" href="styles.css">

<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>

<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>


</head>
<body  topmargin="0" leftmargin="0" style="text-align: center">
<table border="0" cellpadding="0" cellspacing="0" width="657" id="table1" style="font-weight: bold">
<tr><td valign="top"><div align="center">
<font color="#000080" size="6"><%=Request.QueryString("RecName")%></font><br><br></bt><%=Age(Request.QueryString("Lo"),Request.QueryString("Hi"))%>
<table border="0" cellpadding="0" cellspacing="0" width="667" id="table3" style="font-weight: bold; font-size:10pt">
<tr height ="10"><td width="60">&nbsp;</td><td width="35"></td><td width="62"></td>
<td width="75"></td><td width="367"></td><td width="58"></td><td width="11"></td></tr><%while(!rs.eof)
{
if(gender != rs.Fields("Sex").Value)
{gender = rs.Fields("Sex").Value;%><tr><td colspan ="7" algin="center" height="30" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: dashed; border-bottom-width: 1px;"><font color="#000080" size="3"><%if(rs.Fields("Sex") == "M"){%>Male<%}else{if(rs.Fields("Sex") == "F"){%>Female<%}else{%>Mixed<%}}%></font></td></tr>
<%}%><tr valign ="top" <%if(rs.Fields("ago")<=12 & rs.Fields("ago") > 0){%>style="background-color: #6699FF"<%}%>><td><%=CTime(rs.Fields("RecTime"))%></td><td><%=rs.Fields("Distance")%></td><td><%=strokes[rs.Fields("Stroke")]%></td><td><%=rs.Fields("RecDate")%></td><td><%=rs.Fields("RecText")%></td><td><%=rs.Fields("RecTeam")%></td><td><%=rs.Fields("Ex")%></td></tr>
<% rs.MoveNext();}%></table><%
}}rs.Close();oConn.Close();
%></div><p>&nbsp;</td></tr></table>
	</table>
	<p><br>
&nbsp;</p>

</table>

</body></html>