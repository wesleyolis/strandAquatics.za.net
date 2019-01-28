<%
Set objErrMail= Server.CreateObject("CDO.Message")
With objErrMail
.From = "Rupert <ru@stevestats.com>"
.To = "wesleyo@telkomsa.net"
.Subject = "Subject Line"
.HTMLBody = CStr("hello")
.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.strandaquatics.za.net"
.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
.Configuration.Fields.Update
.Send
End With
%>
Email should be sent if no errors 