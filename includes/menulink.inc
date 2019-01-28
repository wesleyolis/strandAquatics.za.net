<script language="javascript">
var url = document.location.href,urldepth="";
l = url.split("/").length;
for(i=6;i<l;i++)
{urldepth += "../";}
	
document.writeln("<link rel='stylesheet' type='text\/css' href='" + urldepth + "styles.css'>");

<!--#include file="menu_depth.inc"-->

</script>