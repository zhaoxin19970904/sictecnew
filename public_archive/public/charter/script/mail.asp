<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Bus Charter</title>
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
e_mail=request.form("e_mail")
book_date=request.form("book_date")
book_way=request.form("book_way")
book_car=request.form("car7")&" "&request.form("car15")&" "&request.form("car21")&" "&request.form("car25")&" "&request.form("car32")
other_require=request.form("other_require")
body="<div align='center'><font color='#ff9900' size=4>Bus Charter</font></div>"&"<br>"
body=body&"Name:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_name&"</font>"&"<br>"
body=body&"Organization:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_unit&"</font>"&"<br>"
body=body&"Title:"&"&nbsp;&nbsp;<font color='#cc0099'>"&buyer_place&"</font>"&"<br>"
body=body&"Phone:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_tel&"</font>"&"<br>"
body=body&"E-mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br>"
body=body&"Date:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_date&"</font>"&"<br>"
body=body&"Route:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_way&"</font>"&"<br>"
body=body&"Vehicle:"&"&nbsp;&nbsp;<font color='#cc0099'>"&book_car&"</font>"&"<br>"
body=body&"Other requirement:"&"&nbsp;&nbsp;<font color='#cc0099'>"&other_require&"</font>"

sql="select * from e_mail where ifuse='4'"
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
 JMail.Charset="gb2312"    
 JMail.From=""&email&""    
 JMail.fromname=""&e_mail&""            
 JMail.mailserverusername=""&login_name&""   
 JMail.MailserverPassword=""&login_pass&""      
 JMail.addRecipient ""&email&""    
 JMail.Subject=""&book_name&"  包车  "&book_car&"  cars"&""    
 JMail.Body=""&body&""
 ifsend=JMail.send (""&server_smtp&"")    
 JMail.close
 set JMail=nothing
 if ifsend=true then
    send="Your request has been submitted, we will contact you soon.Thanks!"
  else
	send="Time out. Pls try again!"
 end if
%>
<script language="javascript">
   alert("<%=send%>")
   history.back()
</script>
<%end if%>
</body>
</html>
