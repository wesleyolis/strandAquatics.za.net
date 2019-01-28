
<%@ Language=JavaScript %>
<%
Response.CacheControl = "Public";
Response.Expires = 2880;

%>
l="";
my = 100;
var tlayout = " Selectedindex='-1'  imenu='1'  border='2' cellpadding='0' cellspacing='0' bordercolor='#000080' style='border-collapse: collapse; position:absolute; left:252px; top:116px; z-index:1'  class='menu'>";



<%
var prov = new Array("WP","BO","CG","EP","ES","FS","GW","KZ","MP","NF","NN","NP","NT","NW","VT");
var oConn;	
	var rs;		
	var filePath;	
	filePath = Server.MapPath("Swim.mdb");
	oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
	for(i=0;i<1;i++)
	{
%>var prov<%=(i+1)%> = new Array(<%
	rs = oConn.Execute("SELECT TEAM.Team, TEAM.TCode, TEAM.TName FROM TEAM WHERE (((TEAM.LSC)='" + prov[i] + "')) ORDER BY TEAM.TName;");
while(!rs.eof)
{
Response.Write("'" + Server.HTMLEncode(rs.Fields("TName").Value) + "','" + Server.HTMLEncode(rs.Fields("TCode").Value) + "','" + Server.HTMLEncode(rs.Fields("Team").Value)+"'");
rs.MoveNext();
if(!rs.eof)
{Response.Write(",");}
}
%>);




<%}%>

var alpha = new Array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
var m0m3 = new Array('Time LCM - 50m','L','Time SCM - 25m','S');
var m0m3m1 = new Array('All Province\'s','All','Border','BO','Central Guateng ','CG','Eastern Province','EP','Eastern Swimming','ES','Freestate','FS','Griqualand West ','GW','Kwa-Zulu Natal','KZ','Mpumalanga ','MP','Northern Freestate','NF','Northern Natal','NN','Northern Province','NP','Northern Tigers','NT','North West','NW','Vaal Triangle','VT','Western Province','WP');
var m0m3m1m1 = new Array('Open','&Lo=0&Hi=99','21 &amp; Over','&Lo=21&Hi=99','18 - 21','&Lo=18&Hi=21','19 & over','&Lo=19&Hi=99','17 - 18','&Lo=17&Hi=18','17 & Over','&Lo=17&Hi=99','16-17','&Lo=16&Hi=17','15 - 18','&Lo15=&Hi=18','15-16','&Lo=15&Hi=16','15 & Over','&Lo=15&Hi=99','13 - 14','&Lo=13&Hi=14','13 - 16','&Lo=13&Hi=16','11- 12','&Lo=11&Hi=12','11 -14','&Lo=11&Hi=14','19 yrs','&Lo=19&Hi=19','18 yrs','&Lo=18&Hi=18','17 yrs','&Lo=17&Hi=17','16 yrs','&Lo=16&Hi=16','15 yrs','&Lo=15&Hi=15','14 yrs','&Lo=14&Hi=14','13 yrs','&Lo=13&Hi=13','12 yrs','&Lo=12&Hi=12','12 & under','&Lo=0&Hi=12','11 yrs','&Lo=11&Hi=11','10 yrs','&Lo=10&Hi=10','10 & under','&Lo=0&Hi=10','9 yrs ','&Lo=9&Hi=9','8 & under ','&Lo=0&Hi=8');
var m0m3m1m1m1 = new Array('50m Freestyle','&Stroke=1&Distance=50&Gender=F','100m Freestyle','&Stroke=1&Distance=100&Gender=F','200m Freestyle','&Stroke=1&Distance=200&Gender=F','400m Freestyle','&Stroke=1&Distance=400&Gender=F','800m Freestyle','&Stroke=1&Distance=800&Gender=F','1500m Freestyle','&Stroke=1&Distance=1500&Gender=F','50m Back','&Stroke=2&Distance=50&Gender=F','100m Back','&Stroke=2&Distance=100&Gender=F','200m Back','&Stroke=2&Distance=200&Gender=F','50m Breast','&Stroke=3&Distance=50&Gender=F','100m Breast','&Stroke=3&Distance=100&Gender=F','200m Breast','&Stroke=3&Distance=200&Gender=F','50m Butterfly','&Stroke=4&Distance=50&Gender=F','100m Butterfly','&Stroke=4&Distance=100&Gender=F','200m Butterfly','&Stroke=4&Distance=200&Gender=F','200m Medley','&Stroke=5&Distance=200&Gender=F','400m Medley','&Stroke=5&Distance=400&Gender=F','50m Freestyle','&Stroke=1&Distance=50&Gender=M','100m Freestyle','&Stroke=1&Distance=100&Gender=M','200m Freestyle','&Stroke=1&Distance=200&Gender=M','400m Freestyle','&Stroke=1&Distance=400&Gender=M','800m Freestyle','&Stroke=1&Distance=800&Gender=M','1500m Freestyle','&Stroke=1&Distance=1500&Gender=M','50m Back','&Stroke=2&Distance=50&Gender=M','100m Back','&Stroke=2&Distance=100&Gender=M','200m Back','&Stroke=2&Distance=200&Gender=M','50m Breast','&Stroke=3&Distance=50&Gender=M','100m Breast','&Stroke=3&Distance=100&Gender=M','200m Breast','&Stroke=3&Distance=200&Gender=M','50m Butterfly','&Stroke=4&Distance=50&Gender=M','100m Butterfly','&Stroke=4&Distance=100&Gender=M','200m Butterfly','&Stroke=4&Distance=200&Gender=M','200m Medley','&Stroke=5&Distance=200&Gender=M','400m Medley','&Stroke=5&Distance=400&Gender=M');
var m0m12 = new Array("All Courses","&Course=A&Gender=F","Long Couse 50m","&Course=L&Gender=F","Short Course 25m","&Course=S&Gender=F","All Courses","&Course=A&Gender=M","Long Couse 50m","&Course=L&Gender=M","Short Course 25m","&Course=S&Gender=M");
makemenu(depth,urldepth);
function makemenu(dep)
{

for(d=0;d<dep;d++)
{
l+="../";
}

if(document.documentElement){
<%
for(i2=0;i2<1;i2++)
	{%>
creatmenu(144,my + 115,'sub','m0','m0m4',2,new Array(245,45,0),18,2,1,prov<%=(i2+1)%>,"",new Array(),"Clubs");
<%}%>
creatmenu(724,my + 115,'link','m0m4','m0m4m1',1,new Array(50,0),18,2,0,alpha,"parent.self.location.replace(\'"+l+"times.asp?TEAM=\' + prov1[(document.getElementById(\'m0m4\').Selectedindex*3)-1] + \'&TName=\' + prov1[(document.getElementById(\'m0m4\').Selectedindex*3)-3] + \'#\' + alpha[(document.getElementById(\'m0m4m1\').Selectedindex)-1]);",new Array(),"Alpha");

creatmenu(144,my + 77,'sub','m0','m0m3',1,new Array(110,0),18,1,1,m0m3,'');
//creatmenu(252,my + 77,'sub','m0m3','m0m3m1',2,new Array(171,30),18,1,0,m0m3m1,'');

//392,116
creatmenu(252,my + 77,'sub','m0m3','m0m3m1',1,new Array(90,0),18,2,1,m0m3m1m1,'',new Array(),"Age Group");
//572,116
creatmenu(432,my + 77,'link','m0m3m1','m0m3m1m1',1,new Array(125,0),18,2,1,m0m3m1m1m1,"parent.self.location.replace(\'"+l+"rankings.asp?Province=WP&Course=\' + m0m3[(document.getElementById(\'m0m3\').Selectedindex*2)-1] + m0m3m1m1[(document.getElementById(\'m0m3m1\').Selectedindex*2)-1] + m0m3m1m1m1[(document.getElementById(\'m0m3m1m1\').Selectedindex*2)-1] + \'&Page=1\');",new Array('Female','Male'));

creatmenu(144,my + 96,'sub','m0','m0m12',1,new Array(125,0),18,2,1,m0m12,"",new Array('Female','Male'));
creatmenu(394,my + 96,'link','m0m12','m0m12m1',1,new Array(90,0),18,2,1,m0m3m1m1,"parent.self.location.replace(\'"+l+"points.asp?Province=WP\' + m0m12[(document.getElementById(\'m0m12\').Selectedindex*2)-1] + m0m3m1m1[(document.getElementById(\'m0m12m1\').Selectedindex*2)-1]);",new Array(),"Age Group");

}

//compatible with versions
v="";
if(document.documentElement){
v +="<tr onmouseover=\"o('m0',3);\" onmouseout=\"t('m0',3)\" onclick=\"sm('m0',3)\" id=\"im03\"><td height=\"19\">Rankings</td></tr>"
v +="<tr onmouseover=\"o('m0',12)\" onmouseout=\"t('m0',12)\" onClick=\"sm('m0',12)\" id=\"im012\"><td height=\"19\">Point Rankings</td></tr>"
v +="<tr onmouseover=\"o('m0',4)\" onmouseout=\"t('m0',4)\" onClick=\"sm('m0',4)\" id=\"im04\"><td height=\"19\">Athlete Times</td></tr>"
}else{
v +="<tr onmouseover=\"o('m0',3);\" onmouseout=\"t('m0',3)\" onclick=\"parent.self.location.replace('"+l+"rankings.asp');\" id=\"im03\"><td height=\"19\">Rankings</td></tr>"
v +="<tr onmouseover=\"o('m0',12)\" onmouseout=\"t('m0',12)\" onClick=\"parent.self.location.replace('"+l+"points.asp');\" id=\"im012\"><td height=\"19\">Point Rankings</td></tr>"
v +="<tr onmouseover=\"o('m0',4)\" onmouseout=\"t('m0',4)\" onClick=\"parent.self.location.replace('"+l+"times.asp');\" id=\"im04\"><td height=\"19\">Athlete Times</td></tr>"
}

document.writeln(""
+"<div align=\"left\">"
+"<table align='left' border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"156\" id=\"table1\" height=\"510\"><tr><td height=\"4\" align=\"center\" valign=\"top\"></td></tr><tr>"
+"<td align=\"center\" valign=\"top\" height=\"27\"><img border=0 width=145 height=26 src=\""+l+"images/strand.gif\" alt=\"Strand Aquatics\"></td></tr>"
+"<tr>"
+"<td align=\"center\" valign=\"bottom\" height=\"50\"><img border=\"0\" src=\""+l+"images/swimmer.gif\" width=\"129\" height=\"39px\"></td>"
+"</tr><tr><td height=\"6\" align=\"center\" valign=\"top\"></td></tr>"
+"<tr><td align=\"center\" valign=\"top\" height=\"305px\">"
+"<table main=\"\" border=\"2\" cellpadding=\"0\" cellspacing=\"0\" width=\"136\" id=\"m0\" bordercolor=\"#000080\" style=\"border-collapse: collapse;display:block; padding-left:3px; position:absolute; left:10px; top:"+my+"px; z-index:1\" class=\"menu\">"
+"<tr onmouseover=\"o('m0',1);\" onmouseout=\"t('m0',1)\" onClick=\"parent.self.location.replace('"+l+"default.asp');\" id=\"im01\"><td height=\"19\">Home</td></tr>"
+"<tr onmouseover=\"o('m0',14);\" onmouseout=\"t('m0',14)\" onClick=\"parent.self.location.replace('"+l+"news/');\" id=\"im014\"><td height=\"19\">W.P News</td></tr>"
+"<tr onmouseover=\"o('m0',15);\" onmouseout=\"t('m0',15)\" onClick=\"parent.self.location.replace('"+l+"meets/meet_info.asp');\" id=\"im015\"><td height=\"19\">Meet Info</td></tr>"
+"<tr onmouseover=\"o('m0',2);\" onmouseout=\"t('m0',2)\" onClick=\"parent.self.location.replace('"+l+"meets/meet.asp');\" id=\"im02\"><td height=\"19\">Meet Results</td></tr>"
+ v
+"<tr onmouseover=\"o('m0',5)\" onmouseout=\"t('m0',5)\" onClick=\"parent.self.location.replace('"+l+"records.asp');\" id=\"im05\"><td height=\"19\">Record Archives</td></tr>"
+"<tr onmouseover=\"o('m0',6)\" onmouseout=\"t('m0',6)\" onClick=\"parent.self.location.replace('"+l+"standards.asp');\" id=\"im06\"><td height=\"19\">Time Standards</td></tr>"
+"<tr onmouseover=\"o('m0',7)\" onmouseout=\"t('m0',7)\" onClick=\"parent.self.location.replace('"+l+"wintergala/index.htm');\" id=\"im07\"><td height=\"19\">Winter Gala</td></tr>"
+"<tr onmouseover=\"o('m0',11)\" onmouseout=\"t('m0',11)\" onClick=\"parent.self.location.replace('"+l+"conduct.htm');\" id=\"im011\"><td height=\"19\">Code of Conduct</td></tr>"
+"<tr onmouseover=\"o('m0',8)\" onmouseout=\"t('m0',8)\" id=\"im08\" onClick=\"parent.self.location.replace('"+l+"western_province/default.asp');\"><td height=\"19\">W.P Login</td></tr>"
+"<tr onmouseover=\"o('m0',13);\" onmouseout=\"t('m0',13)\" onclick=\"parent.self.location.replace('"+l+"subscribe.asp');\" id=\"im013\"><td height=\"19\">E-Mail Notify</td></tr>"

+"<tr  onmouseover=\"o('m0',9)\" onmouseout=\"t('m0',9)\"  onClick=\"parent.self.location.replace('"+l+"contact.htm');\" id=\"im09\"><td height=\"19\">Contact us</td></tr>"
+"<tr  onmouseover=\"o('m0',10)\" onmouseout=\"t('m0',10)\" onClick=\"parent.self.location.replace('"+l+"content.htm');\" id=\"im010\"><td height=\"19\">Terms&nbsp; of use</td>"
+"</tr></table></td></tr><tr><td Style='padding-left: 20;'><a href='mailto:wesley.olis@gmail.com?subject=Wesley Innovative Programming'><img border='0' src='"+l+"images/design.gif' width='96' height='78'></a></td></tr></table></div>");

}
