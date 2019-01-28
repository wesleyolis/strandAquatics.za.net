function GetXmlHttpObject(handler)
{ 
var objXmlHttp;
if (navigator.userAgent.indexOf("Opera")>=0)
   {
    alert("This example doesn't work in Opera") 
    return  
   }
if (navigator.userAgent.indexOf("MSIE")>=0)
   { 
   var strName="Msxml2.XMLHTTP"
   if (navigator.appVersion.indexOf("MSIE 5.5")>=0)
      {
      strName="Microsoft.XMLHTTP"
      } 
   try
      { 
      objXmlHttp=new ActiveXObject(strName)
      return objXmlHttp
      } 
   catch(e)
      { 
      alert("Error. Scripting for ActiveX might be disabled") 
      return 
      } 
    } 
if (navigator.userAgent.indexOf("Mozilla")>=0)
   {
   objXmlHttp=new XMLHttpRequest()
   return objXmlHttp
   }
} 

function loading(msg)
{window.status = "Loading HTML URL " + msg;}
function loaded(msg)
{window.status="";}



function writeHtmlUrl(url,pageobject)
{var xmlHttp;
pageobject=(pageobject!=null)?pageobject:document;
xmlHttp=GetXmlHttpObject();
xmlHttp.onreadystatechange=function() {if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
{if(xmlHttp.status==200){pageobject.writeln(xmlHttp.responseText);
loaded();}else{alert("Error, Code: "+ xmlHttp.status+"URL: "+url);}}}

xmlHttp.open("GET",url, true);
loading();
xmlHttp.send(null);
}

