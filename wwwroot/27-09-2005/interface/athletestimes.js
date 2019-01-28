var strokes=new Array("","Freestyle","Back","Breast","Butterfly","Medley");
var rowset=10;
var CR=0, STR=1, DIS=2, NT=3, SR=4, FP=5, IR=6, AGE=7, MT=8, MTE=9, MN=10;
var display="";

function MTF(data,p)
{return "<a href='../" + data[p+=MT] + "" + data[p++] + "'>" + data[p++] + "</a>";}




function wrt(str)
{display+=str;
if(display.length>125)
{
document.write(display);
display="";
}
}

function wrtFlush()
{document.write(display);}

function wrt_r(str){
wrt("<td>"+str+"</td>");}

function Blank_ArrayCopy2D(R,C)
{
if(C!=null)
for(g=0;g<C.length;g++)
if(C[g].length> 0)
{R[g]=new Array();
for(g2=0;g2<C[g].length;g2++)
R[g][g2]=null;
}else{R[g]=null;}
}

function grouping()
{
code="";
}

function DisAthTimes(data_ath,grp)
{
var Tgrp=new Array();

if(grp!=null)
Blank_ArrayCopy2D(Tgrp,grp)
var stime = new Date;
var glevel=-1;
wrt("<Table border='0'>")
for(i=0;i<data_ath.length;i++)
{
if(grp!=null)
for(g=0;g<grp.length;g++)
if(grp[g].length>0)
{
	
	grps=grp[g];
	for(g2=0;g2<grps.length;g2++)
	{
		if(Tgrp[g][g2]!=data_ath[i][grps[g2]])
		{if(glevel>=g){for(l=(grp.length-g);l>0;--l)wrt("</Table></td></tr>");
		//footer inset here
		}

		glevel=g;
		wrt("<tr><td bgColor='#99CCFF' onclick=\"disgrp('"+g+"_"+i+"');\" colspan='"+ rowset +"'>");
		for(g3=0;g3<grps.length;g3++)
		{
		Tgrp[g][g3]=data_ath[i][grps[g3]];
		wrt(" - " + data_ath[i][grps[g3]]);
		}
		wrt("</td></tr>");
		dis =(g==0)?"block":"none";
		wrt("<tr><td><Table style='display:"+dis +";' border='0' id='grp"+g+"_"+i+"'>");
		g2 = grps.length +1;
		}	
	}
}
else if(Tgrp[g]!=data_ath[i][grp[g]])
	{
	wrt("<tr><td colspan='"+ rowset +"'>");
	Tgrp[g]=data_ath[i][grp[g]];
	wrt(data_ath[i][grp[g]]);
	wrt("</td></tr>");
	}



temp="<tr>";
for(i2=0;i2<rowset;i2++)
temp += "<td>"+ data_ath[i][i2] +"</td>";
wrt(temp+"</tr>");
}
if(grp!=null)
wrt("</Table>")
wrt("</Table>")
endt = new Date;
alert((endt - stime));
wrtFlush();
}


function getjs(urls,i) //downloads js file and excute's it registering all of its functions
{alert('df2sdsfsdfdf');
var url=urls[i];
var xmlHttp;
alert(url+";"+i);
i++;
if(url!=null){

xmlHttp=GetXmlHttpObject();
xmlHttp.onreadystatechange=function() {if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
{if(xmlHttp.status==200){
temp = xmlHttp.responseText;
alert(temp);
eval("alert('eval');"+temp);
Score(2,2066)

getjs(urls,i)
}else{alert("Error, Code: "+ xmlHttp.status+"URL: "+url);}}}
xmlHttp.open("GET",url, true);
xmlHttp.send(null);
}
}
var lastdisgrp;
function disgrp(g)
{
//disgrp2(lastdisgrp,false);
lastdisgrp=g;
setTimeout("disgrp2(lastdisgrp,true);",0);

}

function disgrp2(g,o)
{e=document.getElementById("grp"+g);
if(true){
if(e.style.display !='block')
e.style.display='block';
else
e.style.display='none';}
}
