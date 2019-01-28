
function creatmenu(x,y,type,main,id,tcol,tcwidth,theight,coloms,value,options,code,chead,heading)
{

end = ((options.length%coloms)*(tcol+value));
for(i=0;i<end;i++)
{
options[options.length]='';
}


document.writeln("<table  Selectedindex='-1' main ='" + main + "' id='"+id+"'");
if(type=='link')
{
document.writeln(" onclick=\"stm('"+id+"');" + code + "\"");
}
if(type=='sub')
{
document.writeln(" onclick=\"stm('"+id+"');\" ");
}
document.writeln("cellpadding='0' cellspacing='0' style='border: 1px solid #000080; padding-left: 0; padding-right: 0; padding-top: 0; padding-bottom: 0px;position:absolute; left:"+x+"px; top:"+y+"px; z-index:1;font-size:10pt;background:#B3C6EC; color:#000080; font-family:Arial Rounded MT Bold' class='menu'>");	

	var ttcwidth = 0;
	for(i=0;i<tcwidth.length-1;i++)
	{ttcwidth += tcwidth[i];}
	if(heading != null && heading !="")
		{
		document.writeln("<tr bgcolor='#97B0E6' align='center' height='"+(theight+3)+"px'><td colspan='"+coloms+"' style='border-left-style: solid; border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px'>" + heading + "</td></tr>");
		}
		document.writeln("<tr>");
	for(col = 0; col <coloms;col++)
	{
	document.writeln("<td width='" + ttcwidth + "'>");
	document.writeln("<table cellpadding='0' cellspacing='0'  border='1' bordercolor='#000080' style='border-collapse: collapse;font-size:10pt; padding-left:3px;background:#B3C6EC; color:#000080; font-family:Arial Rounded MT Bold' class='menu2'>");
		if(chead != null && chead.length > 0)
		{
		document.writeln("<tr bgcolor='#97B0E6' align='center' height='"+(theight+2)+"px'><td>" + chead[col] + "</td></tr>");
		}
		for(i=(col*(options.length/coloms)),e=((col)*((options.length/(tcol+value))/(coloms)))+1;i<((col+1)*(options.length/coloms));i=i+tcol+value,e++)
		{if(options[i] != '')
		{
		document.writeln("<tr onmouseover=\"o('"+id+"'," + e +")\" onmouseout=\"t('"+id+"'," + e +")\" onclick=\"s('"+id+"'," + e +")\" id='i"+id + e +"' height='"+theight+"px'>");
			}else
			{
			document.writeln("<tr height='"+theight+"px'>");
		
			}
	for(c=0;c<tcol;c++)
			{
			document.writeln("<td width='"+tcwidth[c]+"'>" + options[i+c] + "</td>");
			}

		document.writeln("</tr>");
		}
	
	document.writeln("</table></td>");
	}
document.writeln("</tr></Table>");
}

var overcol = "6389E2",selCol = "6389F8"
var lastmenu;
lastmenu = '';
function setlm(lastm)
{
//if(lastmenu.indexOf(lastm)>0)
//{
lastmenu = lastm;
//}
}

function o(menu,row)
{
document.getElementById('i' + menu + '' + row).style.backgroundColor = overcol;}

function t(menu,row)
{
if(document.getElementById(menu).Selectedindex!= row)
{
document.getElementById('i' + menu + '' + row).style.backgroundColor = '';}else{
document.getElementById('i' + menu + '' + row).style.backgroundColor = selCol;
}}

function s(menu,row)
{
document.getElementById('i' + menu + '' + row).style.backgroundColor = '6389D8';
document.getElementById(menu).Select = row;
}

function stm(menu)
{
sel = document.getElementById(menu).Select;
if(sel == document.getElementById(menu).Selectedindex)
{sc(lastmenu,menu);
document.getElementById(menu).Selectedindex = -1;
}else{
if(document.getElementById(menu).Selectedindex > -1)
{
document.getElementById('i' + menu + '' + document.getElementById(menu).Selectedindex).style.backgroundColor = '';
}
document.getElementById(menu).Selectedindex = sel;
document.getElementById('i' + menu + '' + sel).style.backgroundColor = '6389D8';
if(document.getElementById(menu+'m1') != null){
document.getElementById(menu+'m1').style.display='block';
document.getElementById(menu+'m1').focus();
}
setlm(menu+'m1');
}
}

function sm(menu,row)
{
sel = document.getElementById(menu).Selectedindex;
if(sel == row)
{sc(lastmenu,menu,true);
document.getElementById('i' + menu + '' + row).style.backgroundColor = '6389D8'
document.getElementById(menu).Selectedindex = -1;
}else{
if(sel > -1)
{

sc(lastmenu,menu);
document.getElementById('i' + menu + '' + sel).style.backgroundColor = '';
}
document.getElementById(menu).Selectedindex = row;
document.getElementById('i' + menu + '' + row).style.backgroundColor = '6389D8';
document.getElementById(menu+'m'+row).style.display='block';
document.getElementById(menu+'m'+row).focus();
setlm(menu+'m'+row);
}
}


function sc(menu,stop)
{
while(menu != stop)
{
sel = document.getElementById(menu).Selectedindex;
if(sel > -1)
{
document.getElementById('i' + menu + '' + sel).style.backgroundColor = '';
document.getElementById(menu).Selectedindex = -1;
}
if(document.getElementById(menu).main !="")
{document.getElementById(menu).style.display='none';}
menu = document.getElementById(menu).main;
}
setlm(stop);
}
