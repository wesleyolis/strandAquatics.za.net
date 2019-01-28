

function Time (Score)
{if(Score==0)
{return '';}
else
{t=(Score<1000)?('0'+(Score/100)):(Score<6000)?(Score/100):(Score%100);
if(Score >6000)
{s = ((Score-t)%6000)/100;
m = (Score-t-(s*100))/6000;
s=(s<10)?'0'+s:s;
m=(m<10)?'0'+m:m;
t= m+':'+s+'.'+t;}
return t;}};

function Age(Lo,Hi)
{
return (Lo>=Hi)?Hi+'yrs':(Lo>0&Hi<99)?Lo+' - '+Hi+'yrs':(Lo==0&Hi==99)?'Open':(Hi==99)?Lo+'&nbsp;&amp;&nbsp;Over':Hi+'&nbsp;&amp;&nbsp;Under';
};
/*
function Score(nt,s)
{if(nt==2)
return 'D\Q (' + Time(s) + ')';
else
return Time(s);
};*/