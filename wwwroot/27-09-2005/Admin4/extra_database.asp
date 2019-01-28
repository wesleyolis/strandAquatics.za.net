<%@ Language=VBScript %>
<%
If Request.Form("DBName").count > 0 And Request.Form("Version").count > 0 And Request.Form("TMV").count > 0 Then
Dim objZip

Response.write "<script language='javascript'><!--/"
Response.write "document.up.WDB.value = 'Extrating Database..';"
Response.write "/--></script>"


Set objZip = Server.CreateObject("XStandard.Zip")
objZip.UnPack "c:\Domains\strandaquatics.com\wwwroot\western_province\secure\Sw" & Request.Form("TMV") & "Bkup" & Request.Form("DBName") & "-" & Request.Form("Version") & ".zip", "c:\Domains\strandaquatics.com\wwwroot\extration"
Response.write "<script language='javascript'><!--/"
Response.write "document.up.WDB.value = 'Extracted';"
Response.write "/--></script>"

Set objZip = Nothing



End IF
%>