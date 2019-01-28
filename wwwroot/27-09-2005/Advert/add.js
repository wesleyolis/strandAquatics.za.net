var i,id,delay;
var images=new Array();
function add_advert(id2,images2,delay2)
{
i=0;
id=id2;
delay = delay2;

for(p=0;p<images2.length;p++)
{	images[p]=new Image();
	images[p].src= "advert/images/" + images2[p];}
roller(0);
}

function roller(i)
{
if(images[i].complete)
{document.getElementById(id).src=images[i].src;
i++;
if(i>=images.length){i=0;}
}setTimeout("roller("+i+")",delay);
}
