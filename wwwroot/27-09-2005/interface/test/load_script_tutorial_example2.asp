<html>
<head>
	<title>Untitled</title>
<style>
DIV,SPAN{position:relative; font-family:arial,helvetica;}
A{font-family:arail,helvetica; text-decoration:none}
A:visited{color:navy}

.cl0{left:0; font-size:15px; height:20}
.cl1{left:10; font-size:13px;}
.cl2{left:10; font-size:12px}
.cl3{left:10; font-size:11px}
.cl4{left:10; font-size:10px}
.clLoad{font-size:13px; font-family:arail,helvetica; color:Silver; left:10}
.clNo{font-size:13px; font-family:arail,helvetica; color:red; left:10}
</style>
<script>
/*********************************************************
Browsercheck
********************************************************/
bw = new Object()
if(document.getElementById && document.createElement && document.appendChild){
  bw.dom = 1
}

/*********************************************************
Vars
********************************************************/
var going=0
var active=new Array()

/*********************************************************
FindElement - this finds the spesified parent element
of the element you send to it. Lazy..nha :}
********************************************************/
function findElement(el,element){
  newel = el
  while(newel.tagName.toLowerCase()!=element && newel.parentNode){
    newel = newel.parentNode
  }
  return newel
}
/*********************************************************
Hide quick
********************************************************/
function hideQuick(obj2){
  obj=document.getElementById(obj2)
  for(j=1;j<obj.childNodes.length;j++){
     obj.childNodes[j].style.display="none"
  }
}
/*********************************************************
Hide all subs of a element
********************************************************/
function hideSubs(o){
  active[o.level]=0
  for(j=0;j<o.childNodes.length;j++){
    if(o.childNodes[j].tagName=="DIV"){
      if(o.childNodes[j].style.display!="") break
      o.childNodes[j].style.display="none"
      if(o.childNodes[j].hasChildNodes()) hideSubs(o.childNodes[j])
    }
  }
}
/*********************************************************
Hiding elements - same effect as showElement
********************************************************/
function hideElements(obj,id,n){
  obj=document.getElementById(obj)
  fn=0
  if(n==-1) n = obj.childNodes.length-1
  if(n>=1){
    obj.childNodes[n].style.display="none"
    n--
    setTimeout("hideElements('"+obj.id+"','"+id+"',"+n+")",50)
  }else if(id!=0){
    active[obj.getAttribute("level")]=0
    going=0
    openMenu(document.getElementById(id))
    fn=1
  }else{ going=0; active[obj.getAttribute("level")]=0; fn=1 }
  
  if(fn){ //Hding all childs
    for(i=(parseInt(obj.getAttribute("level"))+1);i<active.length;i++){
      if(active[i]==0) break;
      else{
        hideQuick(active[i])
        active[i] = 0
      }
    }
  }
}
/*********************************************************
This shows the elements. It uses a timer to make a "cool"
effect when opening the elements
********************************************************/
function showElements(obj,n){
  obj = document.getElementById(obj)
  if(n<obj.childNodes.length){
    obj.childNodes[n].style.display=""
    n++
    setTimeout("showElements('"+obj.id+"',"+n+")",50)
  }else going=0
} 

/*********************************************************
Opening menu
********************************************************/
function openMenu(obj){
  if(!bw.dom) return true;
  obj = findElement(obj,"div")
  if(going) return
  if(obj.getAttribute("level")>active.length) active[obj.getAttribute("level")]=0
  if(active[obj.getAttribute("level")]){ //Another object already open
     going=1
      if(obj.id==active[obj.getAttribute("level")]) id=0
      else id = obj.id
     hideElements(active[obj.getAttribute("level")],id,-1)
  }else{
    going=1
    if(obj.loaded){ //Loaded so we just open it
      active[obj.getAttribute("level")]=obj.id
      showElements(obj.id,1)
    }else{ 
      if(obj.getAttribute("scriptFile")){ //We need to load the content
        div =  document.createElement("div")
        div.id ="loading"
        div.className="clLoad"
        div.appendChild(document.createTextNode("Loading..."))
        obj.appendChild(div)
        activeObj = obj
        loadContent(obj.getAttribute("scriptFile"))
      }else if(obj.link){
        location.href=obj.link
      }else return true
    }
  }
  return false
}
var activeObj;
 
/*********************************************************
Adding elements to the page
********************************************************/
function addElements(txt,scriptfile,link){
  div = document.createElement("div")
  level = (parseInt(activeObj.getAttribute("level")) +1)
  div.className = "cl"+level
  div.setAttribute("level",level)
  div.style.display="none"
  div.id = "divm" + currid++
  span = document.createElement("span")
  span.style.cursor="pointer"
  span.onclick=new Function("openMenu(this)")
  if(scriptfile) div.setAttribute("scriptFile",scriptfile)
  else if(link) div.link = link
  txt = document.createTextNode(txt)
  span.appendChild(txt)
  div.appendChild(span)
  activeObj.appendChild(div)
} 
 
/*********************************************************
This is called when the new elements are loaded
********************************************************/
function loaded(){
  if(activeObj.loaded) return
  activeObj.loaded=1
  going=0
  if(activeObj.childNodes.length==2){
    activeObj.childNodes[1].className="clNo"
    activeObj.childNodes[1].nodeValue="no subs"
  }else{
    activeObj.removeChild(activeObj.childNodes[1]); 
  }
  openMenu(activeObj)
}

/*********************************************************
Loading content from script js file!!
********************************************************/
function loadContent(file){
  var scriptTag = document.getElementById('loadScript');
  var head = document.getElementsByTagName('head').item(0)
  if(scriptTag) head.removeChild(scriptTag);
  script = document.createElement('script');
  script.src = file;
	script.type = 'text/javascript';
	script.id = 'loadScript';
	head.appendChild(script)
}
//Temp to hold the id... uniqueID only works on ie as far as I can tell.
currid = 0
 </script>
 
</head>

<body>
<div level="0" id="m1" scriptFile="item1_subs.js" class="cl0"><a href="LINK_FOR_OLDER_BROWSERS.html" onclick="return openMenu(this)">Item 1</a></div>
<div level="0" id="m2" scriptFile="" class="cl0"><a href="LINK_FOR_OLDER_BROWSERS.html" onclick="return openMenu(this)">Item 2</a></div>

</body>
</html>
