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
else{t = "-";}
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
	if(Request.QueryString("standardFile").count == 0)
	{
	rs = oConn.Execute("SELECT STDNAME.StdFile, STDNAME.Descript, STDNAME.Year FROM STDNAME ORDER BY STDNAME.Year DESC;");
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Time Standards</title>
<!--#include virtual="/includes\headers.inc"-->

<style>
<!--
.r           { border-bottom: 1px solid #0066CC;cursor:hand }
-->
</style>


</head>
<body  topmargin="0" leftmargin="0" style="text-align: center">
<td valign="top"><div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="803"><tr>
<td align="center" valign="top">
<table border="0" cellpadding="0" cellspacing="0" id="table1" style="font-weight: bold">
<tr><td colspan="4" align="center" height="53" valign="bottom"><b>
<font size="6" color="#000080" face="Arial Rounded MT Bold">Time Standards</font></b></td></tr>
<tr style="font-size:8pt;" valign="bottom"><td width="114" height="33">&nbsp;</td>
	<td width="188" height="33">&nbsp;</td><td width="58" height="33">&nbsp;</td>
<td  align="center" width="72" height="33" valign="top"><font color="#000080" size="2"><!--Broken<br>&nbsp;This 
Season--></font></td></tr>
<%while(!rs.eof){%>
<tr><td><a href="standards.asp?standardFile=<%=rs.Fields("StdFile")%>&STDDes=<%=rs.Fields("Descript")%>"><%=rs.Fields("StdFile")%></a></td><td><%=rs.Fields("Descript")%></td><td><%=rs.Fields("Year")%></td><td align="center"></td></tr><%rs.MoveNext()}%></table>
</td>
</td></tr></table>

<%}else{
var recFile = Request.QueryString("standardFile");



if(Request.QueryString("Lo").count == 0 || Request.QueryString("Hi").count == 0 || Request.QueryString("Hi") > 99 || Request.QueryString("Lo") > 99){

//rs = oConn.Execute("SELECT [" + recFile + "].Lo_Age, [" + recFile + "].Hi_Age, Count([" + recFile + "].Rec) AS Broken FROM [" + recFile + "] WHERE ((Year([" + recFile + "]![RecDate])-Year(Date())=0)) GROUP BY [" + recFile + "].Lo_Age, [" + recFile + "].Hi_Age ORDER BY [" + recFile + "].Hi_Age;");
//rs = oConn.Execute("SELECT RECORDS.Lo_Age, RECORDS.Hi_Age, Count(RECORDS.Lo_Age) AS CountOfLo_Age FROM RECORDS WHERE (((((Month(Date()))<5) And ((DateDiff('m',[RecDate],DateSerial((Year(Date())),5,1)))<=12 And (DateDiff('m',[RecDate],DateSerial((Year(Date())),5,1)))>0) Or (((Month(Date()))>=5) And ((DateDiff('m',[RecDate],DateSerial((Year(Date())+1),5,1)))<=12))))) GROUP BY RECORDS.Lo_Age, RECORDS.Hi_Age, RECORDS.Record HAVING (((RECORDS.Record)=" +recFile + "));");


rs = oConn.Execute("SELECT Lo_Age, Hi_age FROM [" + Request.QueryString("standardFile")  + "] GROUP BY Lo_Age,Hi_age ORDER BY Lo_Age;");




var lower, higher;
if(!rs.eof)
{
lower = rs.Fields("Lo_Age") + "";
higher = rs.Fields("Hi_age")+ "" ;
first = "<tr><td></td><td><a href='standards.asp?standardFile=" + recFile + "&STDDes=" + Request.QueryString("STDDes") + "&Lo=" + lower + "&Hi=" + higher + "&Course=LC'>" + Age(lower,higher) + "</td></tr>";
rs.MoveNext();
if(!rs.eof){
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=Request.QueryString("standardFile")%> - Time Standard, <%=Request.QueryString("STDDes")%></title>
<!--#include virtual="/includes\headers.inc"-->

</head>
<body  topmargin="0" leftmargin="0" style="text-align: center">

<font color="#000080" size="6" face="Arial Rounded MT Bold"><%=Request.QueryString("standardFile")%> -  <%=Request.QueryString("STDDes")%></font>

<table border="0" cellpadding="0" cellspacing="0" width="276" id="table2" style="font-weight: bold">
<tr><td align="center" height="43" colspan="2">
<b>Select an Age Group Please</b></td>

</tr><tr height="10"><td width="65" height="8"></td><td width="141" height="8"><a href="standards.asp">Back-Up</a></td>
	</tr>
<%
Response.write(first);															
while(!rs.eof){%><tr><td></td><td><a href="standards.asp?standardFile=<%=recFile%>&STDDes=<%=Request.QueryString("STDDes")%>&Lo=<%=rs.Fields("Lo_Age")%>&Hi=<%=rs.Fields("Hi_Age")%>&Course=LC"><%=Age(rs.Fields("Lo_Age"), rs.Fields("Hi_Age"))%></a></td></tr>
<%rs.MoveNext();
}

}
else
{
Response.Redirect("standards.asp?standardFile=" + recFile + "&STDDes=" + Request.QueryString("STDDes") + "&Lo=" + lower + "&Hi=" + higher + "&Course=LC");
}
}else
{
%>No records found Sorry<%
}%></table><%}else{
var gender,coloms = 1,course = 0,query =" ";
var Course = Request.QueryString("Course");
if(Course == "LC")
{
course = 12;
}else
{
if(Course == "SC")
{
course = 0;//Short
}else
{//if anything else redirect
Response.Redirect("standards.asp?standardFile=" + recFile + "&STDDes=" + Request.QueryString("STDDes") + "&Lo=" + lower + "&Hi=" + higher + "&Course=LC");

}

}

var rst;	
rs = oConn.Execute("SELECT STDNAME.* FROM STDNAME WHERE ((STDNAME.StdFile)='" + Request.QueryString("standardFile") + "');");
%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-za">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=Request.QueryString("standardFile")%> - Time Standard,<%=Age(Request.QueryString("Lo"),Request.QueryString("Hi"))%>, <%=Course%> Course</title>

<!--#include virtual="/includes\headers.inc"-->

</head>
<body  topmargin="0" leftmargin="0" style="text-align: center">
<div align="left">
<table border='0'><tr><td>
<font color="#000080" size="6"><%=Request.QueryString("standardFile")%> -  <%=Request.QueryString("STDDes")%></font><br>
</bt><a href="standards.asp?standardFile=<%=Request.QueryString("standardFile") %>&STDDes=<%=Request.QueryString("STDDes")%>">Back-Up</a><br>
<b>Age Group: <%=Age(Request.QueryString("Lo"),Request.QueryString("Hi"))%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Course:&nbsp;<%if(Course == "LC"){%>Long<%}else{%>Short <%}%>
</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="standards.asp?standardFile=<%=Request.QueryString("standardFile") %>&STDDes=<%=Request.QueryString("STDDes")%>&Lo=<%=Request.QueryString("Lo")%>&Hi=<%=Request.QueryString("Hi")%>&Course=<%if(Course == "LC"){%>SC<%}else{%>LC<%}%>"><b><%if(Course == "LC"){%>Short<%}else{%>Long <%}%> Course</b></a>
<%if(!rs.eof){

for(coloms = 1; coloms <= 12;coloms++){ 
if(rs.Fields("D" + coloms).Value != null && rs.Fields("D" + coloms).Value != "" && rs.Fields("D" + coloms).Value != "    "){
query+= (", IIf(IsNull([S(" + (coloms + course -1) + ")]),0,[S(" + (coloms + course -1) + ")]) As S" + (coloms + course -1));// +  " , ") + ("[S(" + (coloms + course -1)) + ")]";
}else{break;}
}
//Response.write(":" + "SELECT Sex, Distance, Stroke, I_R" + query + " FROM [" + Request.QueryString("standardFile")  + "] WHERE ((Lo_Age = " + Request.QueryString("Lo") + ") And ( Hi_age =" + Request.QueryString("Hi") + ")) ORDER BY Sex,I_R,Stroke,Distance");
rst = oConn.Execute("SELECT Sex, Distance, Stroke, I_R" + query + " FROM [" + Request.QueryString("standardFile")  + "] WHERE ((Lo_Age = " + Request.QueryString("Lo") + ") And ( Hi_age =" + Request.QueryString("Hi") + ")) ORDER BY Sex,I_R,Stroke,Distance;");

%>
<div align="left">
<table><tr><td>
<%while(!rst.eof){
if(rst.fields("Sex").Value != gender)
{gender = rst.fields("Sex").Value;%>
</table><br><table border="0" cellpadding="2" cellspacing="0" style="font-weight: bold; font-size:10pt" bordercolor="#000000"><tr height='25px'><td width='100px'><%if(rst.Fields("Sex") == "M"){%>Male<%}else{%>Female<%}%></td><td width='20px'>I/R</td>
<%
for(coloms = 1; coloms <= 12;coloms++){ 

if(rs.Fields("D" + coloms).Value != null && rs.Fields("D" + coloms).Value != "" && rs.Fields("D" + coloms).Value != "    ")
{%><td width='100px' ><%=rs.Fields("D" + coloms)%></td>
<%}else{break;}
}%>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="font-weight: bold; font-size:10pt" bordercolor="#000000"><tr height='25px'>

<%}%>



<tr><td width='100px'><%=rst.fields("Distance")%>&nbsp;<%=strokes[rst.fields("Stroke")] %></td><td width='20px'><%=rst.fields("I_R")%></td><%

for(c = course; c < (coloms + course - 1);c++){%><td width='100px'><%=CTime(rst.fields("S"+c))%></td>
<%}%></tr>

<%rst.MoveNext();}%>



</table>


</div>
</td></tr>
</table>
</div>
<%
}else
{%>
Sorry No Times found
<%}}}rs.Close(); /*rst.Close();*/oConn.Close(); %>




</body></html>