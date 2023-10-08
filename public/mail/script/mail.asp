<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Charter service</title>
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
   alert("Wrong Email format,right format likes:\nabc123@126.com")
   history.back()
</script>
  <%
response.end()
else
book_name=request.form("book_name")
book_unit=request.form("book_unit")
buyer_place=request.form("buyer_place")
contact_tel=request.form("contact_tel")
book_date=request.form("book_date")
book_way=request.form("book_way")
book_car=request.form("car7")&" "&request.form("car15")&" "&request.form("car21")&" "&request.form("car25")&" "&request.form("car32")
other_require=request.form("other_require")
body="<div align='center'><font color='#ff9900' size=4>Charter</font></div>"&"<br>"
body=body&"Customer's name:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_name&"</font>"&"<br>"
body=body&"Company:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_unit&"</font>"&"<br>"
body=body&"Title:"&"&nbsp;&nbsp;<font color='#cc0099'>"&buyer_place&"</font>"&"<br>"
body=body&"Tel:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_tel&"</font>"&"<br>"
body=body&"Customer's E-mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br>"
body=body&"Charter date:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_date&"</font>"&"<br>"
body=body&"Charter route:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_way&"</font>"&"<br>"
body=body&"Charter cars:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_car&"</font>"&"<br>"
body=body&"Other requirement:"&"&nbsp;&nbsp;<font color='#cc0099'>"&other_require&"</font>"

sql="select * from e_mail where ifuse=true"
 set rs=conn.execute(sql)
if not rs.eof then
  email=rs("email")
  login_name=rs("login_name")
  login_pass=rs("login_pass")
  server_smtp=rs("server_smtp")
  rs.close
 else
  rs.close
end if
Dim JMail,SendMail 
 Set JMail=Server.CreateObject("JMail.Message") 
 JMail.silent=true    
 JMail.Logging=True     
 JMail.ContentType = "text/html" 
 JMail.Charset="utf-8"    
 JMail.From=""&email&""    
 JMail.fromname=""&e_mail&""            
 JMail.mailserverusername=""&login_name&""   
 JMail.MailserverPassword=""&login_pass&""      
 JMail.addRecipient ""&email&""    
 JMail.Subject=""&book_name&"  charter  "&book_car&"  cars"&""    
 JMail.Body=""&body&""
 ifsend=JMail.send (""&server_smtp&"")    
 JMail.close
 set JMail=nothing
 if ifsend=true then
    send="邮件已成功发出."
  else
	send="发送邮件时出错，可能是网络超时."
 end if
%>
<script language="javascript">
   alert("<%=send%>")
   location.href="../index.asp"
</script>
<%end if%>
</body>
</html>
