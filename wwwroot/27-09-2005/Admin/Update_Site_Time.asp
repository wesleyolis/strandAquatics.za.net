<%@ Language=VBScript %>
<%

webdir = "c:/Domains/strandaquatics.com/wwwroot/"
webdir1 = "D:/IIS/wwwroot/27-09-2005/"
Dim todaysDate
todaysDate=now()


if Request.Form("DBName").count > 0 And Request.Form("Version").count > 0 And Request.Form("TMV").count > 0 Then

Response.writeln("<script language='javascript'><!--/" & VBCRLF)
Response.writeln ("document.up.TDis.value = 'Processing..';" & VBCRLF)
Response.writeln("/--></script>" & VBCRLF)
Dim s

dim fs, f 
set fs=Server.CreateObject("Scripting.FileSystemObject") 
set f=fs.CreateTextFile(webdir + "Updated_Date.asp",true) 


f.WriteLine("<%")
f.WriteLine("var DataUpdated = """ & FormatDateTime(todaysDate,2) & """;")
f.WriteLine("var TimeUpdated = """ & FormatDateTime(todaysDate,3) & """;")
f.WriteLine("var DBVersionUpdated = """ & Request.Form("Version") & """;")
f.WriteLine(chr(37) & ">")
f.Close
set f=nothing
set fs=nothing

Response.writeln("<script language='javascript'><!--/" & VBCRLF)
Response.writeln ("document.up.TDis.value = 'Updated';" & VBCRLF)
Response.writeln("/--></script>" & VBCRLF)

else

Response.write("<script language='javascript'><!--/" & VBCRLF)
Response.write ("document.up.TDis.value = 'Error Missing Parameters';" & VBCRLF)
Response.write("/--></script>" & VBCRLF)
End IF%>