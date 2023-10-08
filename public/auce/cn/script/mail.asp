<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Join AUCE</title>
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

mr=request.form("mr")
mrs=request.form("mrs")
miss=request.form("miss")
ms=request.form("ms")
dr=request.form("dr")

txt_cname=request.form("txt_cname")
txt_ename=request.form("txt_ename")
txt_position=request.form("txt_position")
txt_org=request.form("txt_org")
txt_add=request.form("txt_add")
txt_homeadd=request.form("txt_homeadd")
txt_city=request.form("txt_city")
txt_prov=request.form("txt_prov")
txt_code=request.form("txt_code")
txt_country=request.form("txt_country")
txt_workphone=request.form("txt_workphone")
txt_homephone=request.form("txt_homephone")

txt_fax=request.form("txt_fax")
e_mail=request.form("email")
txt_url=request.form("txt_url")

txt_exp=request.form("txt_exp")
txt_other=request.form("txt_other")

individual=request.form("individual")
organizational=request.form("organizational")

yes=request.form("yes")
no=request.form("no")

body="<div align='center'><font color='#ff9900' size=4>Join AUCE</font></div>"&"<br>"

body=body&"Mr."&"&nbsp;&nbsp;<font color='#cc0099'>"&mr&"</font>"&"<br>"
body=body&"Mrs."&"&nbsp;&nbsp;<font color='#cc0099'>"&mrs&"</font>"&"<br>"
body=body&"Miss"&"&nbsp;&nbsp;<font color='#cc0099'>"&fax&"</font>"&"<br>"
body=body&"Ms."&"&nbsp;&nbsp;<font color='#cc0099'>"&ms&"</font>"&"<br><br>"
body=body&"Dr."&"&nbsp;&nbsp;<font color='#cc0099'>"&dr&"</font>"&"<br>"

body=body&"中文名:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_cname&"</font>"&"<br>"
body=body&"英文名:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_ename&"</font>"&"<br>"
body=body&"您的职务:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_position&"</font>"&"<br>"
body=body&"所在团体:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_org&"</font>"&"<br><br>"
body=body&"联系地址:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_add&"</font>"&"<br>"
body=body&"居住地址:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_homeadd&"</font>"&"<br>"
body=body&"所在城市:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_city&"</font>"&"<br>"
body=body&"所在省份:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_prov&"</font>"&"<br>"
body=body&"邮政编码:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_code&"</font>"&"<br><br>"
body=body&"所在国家:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_country&"</font>"&"<br>"

body=body&"办公电话:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_workphone&"</font>"&"<br>"
body=body&"家庭电话:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_homephone&"</font>"&"<br>"
body=body&"传真号码:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_fax&"</font>"&"<br>"
body=body&"电子邮件:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br><br>"
body=body&"您的网址:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_url&"</font>"&"<br>"

body=body&"工作经历:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_exp&"</font>"&"<br>"
body=body&"其他说明:"&"&nbsp;&nbsp;<font color='#cc0099'>"&txt_other&"</font>"&"<br>"

body=body&"个人会员:"&"&nbsp;&nbsp;<font color='#cc0099'>"&individual&"</font>"&"<br>"
body=body&"团体会员:"&"&nbsp;&nbsp;<font color='#cc0099'>"&organizational&"</font>"&"<br>"

body=body&"我同意:"&"&nbsp;&nbsp;<font color='#cc0099'>"&yes&"</font>"&"<br>"
body=body&"不同意:"&"&nbsp;&nbsp;<font color='#cc0099'>"&no&"</font>"&"<br><br>"


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
	send="Time out. Pls try again.Or send your request directly to info@auce.org, thanks!"
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
