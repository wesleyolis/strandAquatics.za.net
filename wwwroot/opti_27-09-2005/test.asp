
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
/*
var strokes = new Array("","Freestyle","Back","Breast","Butterfly","Medley");
	var oConn;	
	var rst;		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	//rst = oConn.Execute("SELECT * FROM RECBREAKERS ORDER BY RecDate");
	
	var rst = oConn.Execute("SELECT RECNAME.RecFile, RECORDS.Record, * FROM RECNAME INNER JOIN RECORDS ON RECNAME.RECORD = RECORDS.Record WHERE (((DateDiff('w',[RecDate],Date()))<=4 And (DateDiff('w',[RecDate],Date()))>=0) AND ((RECORDS.I_R)='I')) ORDER BY RECORDS.RecDate;");
*/
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<title>Strand Aquatics Swimming Club</title>
<meta name="description" content="Strand Aquatic's Swimming Club offers you the change to explorer it's wonder's world ofinformation, such as news, events, results, and the up coming... ">
<meta name="keywords" content="winter,gala,galas,strand,swimming,club,aquatics,WPA,WP,western,province,news,events,results,stran">

<link rel="File-List" href="test_files/filelist.xml">


<link rel="stylesheet" type="text/css" href="styles.css">
<script src="menu.asp" language="javascript" type="Text/Javascript">
</script>
<script language="javascript" type="Text/Javascript">
makemenu(0);
</script>

<style>
<!--
.r           { border-bottom: 1px solid #FF6666;cursor:hand }
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
-->
</style>
<!--[if gte mso 9]>
<xml><o:shapedefaults v:ext="edit" spidmax="1027"/>
</xml><![endif]-->

</head>

<body topmargin="0" leftmargin="0">
<table width="840" height="100%" cellspacing="0" cellpadding="0" id="table1"><tr>
	<td align='center' height="104">
	<table style="position: absolute; left: 169px; top: 49px; width: 540px; height: 51px" border="0" id="Strand2" cellspacing="0" cellpadding="0">
	<tr>
		<td style="display:none;" id="star1" align="center" width="62"><!--[if gte vml 1]><v:shapetype id="_x0000_t136"
 coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">
 <v:formulas>
  <v:f eqn="sum #0 0 10800"/>
  <v:f eqn="prod #0 2 1"/>
  <v:f eqn="sum 21600 0 @1"/>
  <v:f eqn="sum 0 0 @2"/>
  <v:f eqn="sum 21600 0 @3"/>
  <v:f eqn="if @0 @3 0"/>
  <v:f eqn="if @0 21600 @1"/>
  <v:f eqn="if @0 0 @2"/>
  <v:f eqn="if @0 @4 21600"/>
  <v:f eqn="mid @5 @6"/>
  <v:f eqn="mid @8 @5"/>
  <v:f eqn="mid @7 @8"/>
  <v:f eqn="mid @6 @7"/>
  <v:f eqn="sum @6 0 @5"/>
 </v:formulas>
 <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"
  o:connectangles="270,180,90,0"/>
 <v:textpath on="t" fitshape="t"/>
 <v:handles>
  <v:h position="#0,bottomRight" xrange="6629,14971"/>
 </v:handles>
 <o:lock v:ext="edit" text="t" shapetype="t"/>
</v:shapetype><v:shape id="_x0000_s1041" type="#_x0000_t136" alt="*" style='width:37.5pt;
 height:34.5pt' fillcolor="#fff200" stroked="f" strokecolor="#036"
 strokeweight=".1323mm">
 <v:fill color2="#4d0808" rotate="t" angle="-135" colors="0 #fff200;29491f #ff7a00;45875f #ff0300;1 #4d0808"
  method="none" focus="50%" type="gradient"/>
 <v:shadow color="#868686"/>
 <v:textpath style='font-family:"Times New Roman";font-size:28pt;v-text-kern:t'
  trim="t" fitpath="t" string="*"/>
</v:shape><![endif]--><![if !vml]><img border=0 width=51 height=47
src="test_files/image001.gif" alt="*" v:shapes="_x0000_s1041"><![endif]></td>
		<td>
		<p align="center"><!--[if gte vml 1]><v:shape
 id="_x0000_s1040" type="#_x0000_t136" alt="Strand Aquatics Swimming" style='width:309.75pt;
 height:34.5pt' fillcolor="#060" stroked="f" strokecolor="#036" strokeweight=".1323mm">
 <v:fill src="" o:title="Green marble" rotate="t" type="tile"/>
 <v:shadow color="#868686"/>
 <v:textpath style='font-family:"Times New Roman";font-size:28pt;v-text-kern:t'
  trim="t" fitpath="t" string="Strand Aquatics Swimming"/>
</v:shape><![endif]--><![if !vml]><img border=0 width=413 height=46
src="test_files/image002.gif" alt="Strand Aquatics Swimming" v:shapes="_x0000_s1040"><![endif]></td>
		<td id="star2" width="62" align="center" style="display:none;"><!--[if gte vml 1]><v:shape
 id="_x0000_s1039" type="#_x0000_t136" alt="*" style='width:37.5pt;height:34.5pt'
 fillcolor="#fff200" stroked="f" strokecolor="#036" strokeweight=".1323mm">
 <v:fill color2="#4d0808" rotate="t" angle="-135" colors="0 #fff200;29491f #ff7a00;45875f #ff0300;1 #4d0808"
  method="none" focus="50%" type="gradient"/>
 <v:shadow color="#868686"/>
 <v:textpath style='font-family:"Times New Roman";font-size:28pt;v-text-kern:t'
  trim="t" fitpath="t" string="*"/>
</v:shape><![endif]--><![if !vml]><img border=0 width=51 height=47
src="test_files/image001.gif" alt="*" v:shapes="_x0000_s1039"><![endif]></td>
	</tr>
</table>
<table style="position: absolute; left: 190px; top: 0px; width: 502px; height: 57px; z-index: 1" border="0" id="Strand1" cellspacing="0" cellpadding="0">
	<tr>
		<td><!--[if gte vml 1]><v:shape
 id="_x0000_s1038" type="#_x0000_t136" alt="Strand Aquatics Swimming" style='width:375pt;
 height:40.5pt' fillcolor="#add0d3" stroked="f">
 <v:fill opacity="45875f"/>
 <v:shadow color="#868686"/>
 <v:textpath style='font-family:"Times New Roman";font-size:28pt;v-text-kern:t'
  trim="t" fitpath="t" string="Strand Aquatics Swimming"/>
</v:shape><![endif]--><![if !vml]><img border=0 width=500 height=54
src="test_files/image003.gif" alt="Strand Aquatics Swimming" v:shapes="_x0000_s1038"><![endif]></td>
	</tr>
</table>
<script language="javascript" Defer>
<!--
var pos1 = 1,pos2 = 50,way = 1,inner = 200,delay = 10,disp = false,fcount =0;;
setTimeout('move()',2000);

function flash()
{
if(fcount < 6)
{
if(disp == true)
{
disp = false;
document.getElementById("star1").style.display = "none";
document.getElementById("star2").style.display = "none";
setTimeout("flash()",300);
}else{
document.getElementById("star1").style.display = "block";
document.getElementById("star2").style.display = "block";
setTimeout("flash()",500);
disp = true;
}
fcount++;
}
}

function flip(id,dir,wid)
{
document.getElementById(id).style.posWidth -= (dir*20);
var width = document.getElementById(id).style.posWidth;

if(width <= (wid-19))
{
if(width <= inner)
{
document.getElementById(id).style.posWidth =inner+1;
setTimeout("flip('"+id+"',-1,"+wid+")",delay);
}else
{
setTimeout("flip('"+id+"',"+dir+","+wid+")",delay);
}}
else
{
setTimeout("document.getElementById(id).style.posWidth =" + wid + ";",delay);
}


}

function move()
{
pos1 += way;
pos2 -= way;
document.getElementById("Strand1").style.posTop = pos1;
document.getElementById("Strand2").style.posTop = pos2;


if((pos1 > 1 & pos1 < 46))
{
setTimeout('move()',60)
}
else
{
if(pos1 == 1)
{
document.getElementById("Strand1").style.zIndex = 1;
document.getElementById("Strand2").style.zIndex = 0;
}
else
{
document.getElementById("Strand1").style.zIndex = 0;
document.getElementById("Strand2").style.zIndex = 1;

}
	if(way == -1)
	{
	way =1;
	}else
	{
	way =-1;
	}
//setTimeout("flip('_x0000_s1045',1,309);",500);
fcount = 0;
if(pos1 > 1)
{setTimeout("flash()",250);}
setTimeout('move()',3500);
}
}

//-->
</script>

<p>&nbsp;

</td>
	<td align='center' height="104">
&nbsp;</td>
	<td align='center' height="104">
<!--[if gte vml 1]><v:shape
 id="_x0000_s1033" type="#_x0000_t136" alt="Designed&#13;&#10;by&#13;&#10;WIP"
 style='position:absolute;left:550.5pt;top:9pt;width:104.25pt;height:31.5pt;
 z-index:4' fillcolor="#fcc">
 <v:fill src="" o:title="Pink tissue paper" rotate="t" type="tile"/>
 <v:shadow color="#868686"/>
 <v:textpath style='font-family:"Arial Black";font-size:24pt;v-text-kern:t'
  trim="t" fitpath="t" string="Designed"/>
</v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout;position:
absolute;z-index:4;left:733px;top:11px;width:141px;height:44px'><img width=141
height=44 src="test_files/image004.gif" alt="Designed&#13;&#10;by&#13;&#10;WIP"
v:shapes="_x0000_s1033"></span><![endif]><!--[if gte vml 1]><v:shape id="_x0000_s1037"
 type="#_x0000_t136" alt="WIP" style='position:absolute;left:573.75pt;top:37.5pt;
 width:50.25pt;height:25.5pt;z-index:3' fillcolor="#7dbeff">
 <v:fill rotate="t"/>
 <v:shadow color="#868686"/>
 <v:textpath style='font-family:"Times New Roman";font-size:24pt;font-weight:bold;
  v-text-kern:t' trim="t" fitpath="t" string="WIP"/>
</v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout;position:
absolute;z-index:3;left:764px;top:49px;width:69px;height:36px'><img width=69
height=36 src="test_files/image005.gif" alt=WIP v:shapes="_x0000_s1037"></span><![endif]><!--[if gte vml 1]><v:shape
 id="_x0000_s1034" type="#_x0000_t136" alt="by" style='position:absolute;
 left:558.75pt;top:29.25pt;width:18.75pt;height:20.25pt;z-index:5'>
 <v:shadow color="#868686"/>
 <v:textpath style='font-family:"Arial Black";font-size:14pt;v-text-kern:t'
  trim="t" fitpath="t" string="by"/>
</v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout;position:
absolute;z-index:5;left:744px;top:38px;width:28px;height:29px'><img width=28
height=29 src="test_files/image006.gif" alt=by v:shapes="_x0000_s1034"></span><![endif]></td></tr><tr>
		<td align='center' valign='top' width="556">
<!--#include file="news/welcome.asp"-->
	<table border="0" cellpadding="0" cellspacing="0" width="416" id="table7" height="152">
		<tr>
			<td height="5"></td>
		</tr>
		<tr>
			<td height="135" align="center">
			<img border="0" src="images/olmpic_medals.jpg" width="215" height="147" style="border: 3px solid #3366CC"></td>
		</tr>
</table>
</td><td align='center' valign='top' width="4">&nbsp;</td>
		<td align='center' valign='top' width="280">


<%
/*if(!rst.eof){var pos = 0;%>

<table border="4" cellpadding="0" width="274" style="border-collapse: collapse; font-size:8pt" bordercolor="#FF0000" bordercolorlight="#FF6666" bgcolor="#FFBFBF" height="450" onmouseover="over = true;" onmouseout="over=false;">
<tr><td align="center" height="50"><b><font color="#000990" size="4">Congratulations to the following<br>&nbsp; Record Breakers</font></b></td></tr>
<tr><td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="266" id="table8">

<%while(!rst.eof){%>
<tr onmouseout="this.style.backgroundColor='';"  onmouseover="this.style.backgroundColor='#77A9FF';" onClick="self.location.replace('records.asp?recordFile=<%=rst.Fields("RECORDS.Record")%>&RecName=<%=rst.Fields("RecFile")%>&Lo=<%=rst.Fields("Lo_Age")%>&Hi=<%=rst.Fields("Hi_Age")%>');" id="rec<%=pos%>" <%if(pos > 7){%>style="display:none;"<%}%>><td class="r" height="37" valign="top">
<table border="0" cellpadding="2"  width="266" cellspacing="0" style="font-size: 8pt;">
<tr><td  valign="top" align="left" style="font-size: 10pt"><%=Age(rst.Fields("Lo_Age"),rst.Fields("Hi_Age"))%><br><%=CTime(rst.Fields("RecTime"))%></td>
<td  valign="top" width="203"><%=rst.Fields("Distance")%>m&nbsp;<%=strokes[rst.Fields("Stroke")]%>&nbsp;<%=rst.Fields("RecFile")%>&nbsp;Record,<br><b><%=rst.Fields("RecText")%></b>&nbsp;
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
</script><%}*/%>



		</td></tr><tr>
		<td align='center' valign='top' width="840" colspan="3">
<div align="left">
&nbsp;
<script language="javascript">
<!--

function o(menu,row)
{
document.getElementById('i' + menu + '' + row).style.backgroundColor = "6389E2";}

function t(menu,row)
{
document.getElementById('i' + menu + '' + row).style.backgroundColor = '';}

//-->
</script>
<table border="0" id="table16" cellspacing="0" cellpadding="0" style="font-size: 10pt; font-family: Times New Roman; font-weight: bold" width="661">
		<tr>
			<td align="center">

<table class="menu" id="table17" style="display: block; border-collapse: collapse; padding-left: 3px; font-size:10pt; font-family:Times New Roman; font-weight:bold" borderColor="#000080" cellSpacing="0" cellPadding="0" width="647" border="2">
	<tr height="19px">
	<td id="i1m01" onmouseover="o('1m0',1);" onclick="parent.self.location.replace('default.asp');" onmouseout="t('1m0',1)" nowrap><a href="default.asp">
		<font color="#000080">News </font>
	</a></td>
	<td id="i1m02" onmouseover="o('1m0',2);" onclick="parent.self.location.replace('meets/index.htm');" onmouseout="t('1m0',2)" nowrap><a href="meets/index.htm">
		<font color="#000080">Meet Results </font>
	</a></td>
	<td id="i1m03" onmouseover="o('1m0',3);" onclick="parent.self.location.replace('rankings.asp?Province=WP');" onmouseout="t('1m0',3)" nowrap><a href="rankings.asp?Province=WP">
		<font color="#000080">Rankings </font>
	</a></td>
	<td id="i1m012" onmouseover="o('1m0',12)" onclick="parent.self.location.replace('points.asp?Province=WP');" onmouseout="t('1m0',12)" nowrap><a href="points.asp?Province=WP">
		<font color="#000080">Point Rankings </font>
	</a></td>
	<td id="i1m04" onmouseover="o('1m0',4)" onclick="parent.self.location.replace('times.asp?Province=WP');" onmouseout="t('1m0',4)" nowrap><a href="times.asp?Province=WP">
		<font color="#000080">Athlete Times </font>
	</a></td>
	<td id="i1m05" onmouseover="o('1m0',5)" onclick="parent.self.location.replace('records.asp');" onmouseout="t('1m0',5)" nowrap><a href="records.asp">
		<font color="#000080">Record Archives </font>
	</a></td>
	<td id="i1m06" onmouseover="o('1m0',6)" onclick="parent.self.location.replace('qt/index.htm');" onmouseout="t('1m0',6)" nowrap>
	<a href="standards.asp"><font color="#000080">Time Standards </font></a></td>
	</tr>
	</table>
			</td>
		</tr>
		<tr>
			<td align="center">
<table class="menu" id="table18" style="display: block; border-collapse: collapse; padding-left: 3px; font-size:10pt; font-weight:bold; font-family:Times New Roman" borderColor="#000080" cellSpacing="0" cellPadding="0" width="537" border="2">
	<tr height="19px">

	<td id="i1m07" onmouseover="o('1m0',7)" onclick="parent.self.location.replace('wintergala/index.htm');" onmouseout="t('1m0',7)" nowrap>
	<a href="meets/meet_info.asp">Meet Info</a></td>

	<td id="i1m07" onmouseover="o('1m0',7)" onclick="parent.self.location.replace('wintergala/index.htm');" onmouseout="t('1m0',7)" nowrap>
	<a href="news/">W.P Info</a> </td>

	<td id="i1m07" onmouseover="o('1m0',7)" onclick="parent.self.location.replace('wintergala/index.htm');" onmouseout="t('1m0',7)" nowrap><a href="wintergala/index.htm">
		<font color="#000080">Winter Gala </font>
	</a></td>
	<td id="i1m011" onmouseover="o('1m0',11)" onclick="parent.self.location.replace('conduct.htm');" onmouseout="t('1m0',11)" nowrap><a href="conduct.htm">
		<font color="#000080">Code of Conduct </font>
	</a></td>
	<td id="i1m08" onmouseover="o('1m0',8)" onclick="parent.self.location.replace('western_province/default.asp');" onmouseout="t('1m0',8)" nowrap><a href="western_province/default.asp">
		<font color="#000080">W.P Login </font>
	</a></td>
	<td id="i1m013" onmouseover="o('1m0',13);" onclick="parent.self.location.replace('subscribe.asp');" onmouseout="t('1m0',13)" nowrap><a href="subscribe.asp">
		<font color="#000080">E-Mail Notify </font>
	</a></td>
	<td id="i1m09" onmouseover="o('1m0',9)" onclick="parent.self.location.replace('contact.htm');" onmouseout="t('1m0',9)" nowrap><a href="contact.htm">
		<font color="#000080">Contact us </font>
	</a></td>
	<td id="i1m010" onmouseover="o('1m0',10)" onclick="parent.self.location.replace('content.htm');" onmouseout="t('1m0',10)" nowrap><a href="content.htm">
		<font color="#000080">Terms of use </font>
	</a></td></tr>
</table>

			</td>
		</tr>
	</table>
	<br></div>
</td></tr></table>
</td></tr></table></div>
</body>
</html>