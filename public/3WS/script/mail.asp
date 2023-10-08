<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>3WS order</title>
</head>
<%
function IsValidEmail(email)
dim names, name, i, c
IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if
end function
%>
<body>
<%
e_mail=request.form("email")
if IsValidEmail(e_mail)=false then
  %>
 <script language="javascript">
   alert("Wrong Email address,pls check!")
   history.back()
 </script>
  <%
 response.end()
else

contact_name=request.form("contact_name")
contact_tel=request.form("contact_tel")
fax=request.form("fax")
e_mail=request.form("email")
address=request.form("address")

order08=request.form("model0800")
order10=request.form("model1000")
order12=request.form("model1200")
order14=request.form("model1400")

ed_time=request.form("ed_time")
notes=request.form("notes")

body="<div align='center'><font color='#ff9900' size=4>3WS order </font></div>"&"<br>"
body=body&"Name:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_name&"</font>"&"<br>"
body=body&"Tel:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_tel&"</font>"&"<br>"
body=body&"Fax:"&"&nbsp;&nbsp;<font color='#cc0099'>"&fax&"</font>"&"<br>"
body=body&"E-mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br><br>"
body=body&"Address:"&"&nbsp;&nbsp;<font color='#cc0099'>"&address&"</font>"&"<br>"


body=body&"I want to order:"&"&nbsp;&nbsp;<font color='#cc0099'>"&model0800&"</font>"&"<br>"
body=body&"I want to order:"&"&nbsp;&nbsp;<font color='#cc0099'>"&model1000&"</font>"&"<br>"
body=body&"I want to order:"&"&nbsp;&nbsp;<font color='#cc0099'>"&model1200&"</font>"&"<br>"
body=body&"I want to order:"&"&nbsp;&nbsp;<font color='#cc0099'>"&model1400&"</font>"&"<br>"

body=body&"Expected delivery time:"&"&nbsp;&nbsp;<font color='#cc0099'>"&ed_time&"</font>"&"<br>"
body=body&"Notes or comments:"&"&nbsp;&nbsp;<font color='#cc0099'>"&notes&"</font>"&"<br><br>"


sql="select * from e_mail where ifuse='3'"
 set rs=conn.execute(sql)
 if not rs.eof then
  email=rs("email")
  login_name=rs("login_name")
  login_pass=rs("login_pass")
  server_smtp=rs("server_smtp")
  rs.close%>
  <%
  else
   rs.close
 end if
Dim JMail,SendMail 
 Set JMail=Server.CreateObject("JMail.Message") 
 JMail.silent=true     '
 JMail.Logging=True     
 JMail.ContentType = "text/html" 
 JMail.Charset="gb2312"    
 JMail.From=""&e_mail&""    
 JMail.fromname=""&e_mail&""             
 JMail.mailserverusername=""&login_name&""    
 JMail.MailserverPassword=""&login_pass&""     
 JMail.addRecipient ""&email&""     
 JMail.Subject=""&"Ordering 3WS from  "&contact_name&""      
 JMail.Body=""&body&""
 ifsend=JMail.send (""&server_smtp&"")    
 JMail.close
 set JMail=nothing
 if ifsend=true then
    send="Your request has been submitted, we will contact you soon.Thanks!"
  else
	send="Time out. Pls try again.Or send your request directly to trade@sictec.com, thanks!"
 end if
%>
<script language="javascript">
   alert("<%=send%>")
   history.back()
</script>
<%end if%>
<%
  conn.close
  set conn=nothing
  %>

</body>
</html>
