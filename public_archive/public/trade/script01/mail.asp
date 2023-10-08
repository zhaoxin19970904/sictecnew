<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn01.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Sourcing request</title>
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

u_date=request.form("u_date")
subject=request.form("subject")

price=request.form("price")
product=request.form("product")
nearest=request.form("nearest")
packing=request.form("packing")
terms=request.form("terms")
delivery=request.form("delivery")
minimum=request.form("minimum")
discount=request.form("discount")

order=request.form("order")
notes=request.form("notes")

contactperson=request.form("contactperson")
jobtitle=request.form("jobtitle")
tel=request.form("tel")
e_mail=request.form("email")

body="<div align='center'><font color='#ff9900' size=4>Trade Online Inquiry</font></div>"&"<br>"
body=body&"Date:"&"&nbsp;&nbsp;<font color='#cc0099'>"&u_date&"</font>"&"<br>"
body=body&"Subject: "&"&nbsp;&nbsp;<font color='#cc0099'>"&subject&"</font>"&"<br>"

body=body&"You want to know:"&"&nbsp;&nbsp;<font color='#cc0099'>"&price&"</font>"&"<br>"
body=body&"You want to know:"&"&nbsp;&nbsp;<font color='#cc0099'>"&product&"</font>"&"<br>"
body=body&"You want to know:"&"&nbsp;&nbsp;<font color='#cc0099'>"&nearest&"</font>"&"<br>"
body=body&"You want to know:"&"&nbsp;&nbsp;<font color='#cc0099'>"&packing&"</font>"&"<br>"
body=body&"You want to know:"&"&nbsp;&nbsp;<font color='#cc0099'>"&terms&"</font>"&"<br>"
body=body&"You want to know:"&"&nbsp;&nbsp;<font color='#cc0099'>"&delivery&"</font>"&"<br>"
body=body&"You want to know:"&"&nbsp;&nbsp;<font color='#cc0099'>"&minimum&"</font>"&"<br>"
body=body&"You want to know:"&"&nbsp;&nbsp;<font color='#cc0099'>"&discount&"</font>"&"<br>"

body=body&"Expected order volume:"&"&nbsp;&nbsp;<font color='#cc0099'>"&order&"</font>"&"<br>"
body=body&"Further requirements:"&"&nbsp;&nbsp;<font color='#cc0099'>"&notes&"</font>"&"<br>"

body=body&"Contact Person:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contactperson&"</font>"&"<br>"
body=body&"Job Title:"&"&nbsp;&nbsp;<font color='#cc0099'>"&jobtitle&"</font>"&"<br><br>"
body=body&"Telphone: "&"&nbsp;&nbsp;<font color='#cc0099'>"&tel&"</font>"&"<br><br>"
body=body&"E-mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br><br>"

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
 JMail.Subject=""&"Trade Online Inquiry from  "&contactperson&""      
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
