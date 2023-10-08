<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Your China Office</title>
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

B_area=request.form("B_area")
products=request.form("products")
requirements=request.form("requirements")

contact_name=request.form("contact_name")
company=request.form("company")
website=request.form("website")
contact_tel=request.form("contact_tel")
fax=request.form("fax")
e_mail=request.form("email")

body="<div align='center'><font color='#ff9900' size=4>Your China Office</font></div>"&"<br>"
body=body&"Business area:"&"&nbsp;&nbsp;<font color='#cc0099'>"&B_area&"</font>"&"<br>"
body=body&"Products sourcing/purchase:"&"&nbsp;&nbsp;<font color='#cc0099'>"&products&"</font>"&"<br>"
body=body&"Requirements:"&"&nbsp;&nbsp;<font color='#cc0099'>"&requirements&"</font>"&"<br>"

body=body&"Name:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_name&"</font>"&"<br>"
body=body&"Company:"&"&nbsp;&nbsp;<font color='#cc0099'>"&company&"</font>"&"<br>"
body=body&"Website:"&"&nbsp;&nbsp;<font color='#cc0099'>"&website&"</font>"&"<br><br>"
body=body&"Telephone:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_tel&"</font>"&"<br>"
body=body&"Fax:"&"&nbsp;&nbsp;<font color='#cc0099'>"&fax&"</font>"&"<br>"
body=body&"E-mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br><br>"

sql="select * from e_mail where ifuse='2'"
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
 JMail.Subject=""&"Sourcing request from  "&contact_name&""      
 JMail.Body=""&body&""
 ifsend=JMail.send (""&server_smtp&"")    
 JMail.close
 set JMail=nothing
 if ifsend=true then
    send="Your request has been submitted, we will contact you soon.Thanks!"
  else
	send="Time out. Pls try again.Or send your request directly to sourcing@sictec.com, thanks!"
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
