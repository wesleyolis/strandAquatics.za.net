<script language="javascript">


var url = document.location.href, depth = 0, urldepth = "";

i = 0;
i = url.indexOf("www.strandaquatics.za.net/",0) + 26;
i = url.indexOf("/", i);
while(i != -1)
{
	i++;
	depth++;
	urldepth += "../";
	i = url.indexOf("/", i);
}
document.writeln("<link rel='stylesheet' type='text\/css' href='" + urldepth + "styles.css'>");

<!--#include file="menu_depth.inc"-->

</script>