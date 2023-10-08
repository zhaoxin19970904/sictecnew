<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Hotel Booking</title>
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

contact_name=request.form("contact_name")
company=request.form("company")
place=request.form("place")
contact_tel=request.form("contact_tel")
e_mail=request.form("e_mail")

city1=request.form("city1")
city2=request.form("city2")
city3=request.form("city3")

if city1<>"" then
  shop_ask_city1=request.form("shop_ask_city1")
  in_date1=request.Form("in_date1")
  out_date1=request.Form("out_date1")
  shop1_level=request.Form("shop1_level")
  city1_dr=request.Form("city1_dr")
  city1_sr=request.Form("city1_sr")
  city1_tf=request.Form("city1_tf")
  if city1_dr="" then city1_dr="0"
  if city1_sr="" then city1_sr="0"
  if city1_tf="" then city1_tf="0"
end if

if city2<>"" then
  shop_ask_city2=request.form("shop_ask_city2")
  in_date2=request.Form("in_date2")
  out_date2=request.Form("out_date2")
  shop2_level=request.Form("shop2_level")
  city2_dr=request.Form("city2_dr")
  city2_sr=request.Form("city2_sr")
  city2_tf=request.Form("city2_tf")
  if city2_dr="" then city2_dr="0"
  if city2_sr="" then city2_sr="0"
  if city2_tf="" then city2_tf="0"
end if

if city3<>"" then
  shop_ask_city3=request.form("shop_ask_city3")
  in_date3=request.Form("in_date3")
  out_date3=request.Form("out_date3")
  shop3_level=request.Form("shop3_level")
  city3_dr=request.Form("city3_dr")
  city3_sr=request.Form("city3_sr")
  city3_tf=request.Form("city3_tf")
  if city3_dr="" then city3_dr="0"
  if city3_sr="" then city3_sr="0"
  if city3_tf="" then city3_tf="0"
end if

body="<div align='center'><font color='#ff9900' size=4>Hotel Booking</font></div>"&"<br>"
body=body&"Name:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_name&"</font>"&"<br>"
body=body&"Organization:"&"&nbsp;&nbsp;<font color='#cc0099'>"&company&"</font>"&"<br>"
body=body&"Title:"&"&nbsp;&nbsp;<font color='#cc0099'>"&place&"</font>"&"<br>"
body=body&"Phone:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_tel&"</font>"&"<br>"
body=body&"E-Mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br><br>"

if city1<>"" then
body=body&" Booking requirement:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop_ask_city1&"</font>"&"<br>"
body=body&"City #1:"&"&nbsp;&nbsp;<font color='#cc0099'>"&city1&"</font>"&"<br>"
body=body&"Check-in Date: "&"&nbsp;&nbsp;<font color='#cc0099'>"&in_date1&"</font>"&"<br>"
body=body&"Check-out Date: "&"&nbsp;&nbsp;<font color='#cc0099'>"&out_date1&"</font>"&"<br>"
body=body&"Grade:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop1_level&"</font>"&"<br>"
body=body&"No. of rooms:"&"&nbsp;&nbsp;Single:<font color='#cc0099'>"&city1_dr&"</font>room;"
body=body&"&nbsp;&nbsp;Double:<font color='#cc0099'>"&city1_sr&"</font>room;"
body=body&"&nbsp;&nbsp;Suite:<font color='#cc0099'>"&city1_tf&"</font>room<br><br>"
end if

if city2<>"" then
body=body&"City #2:"&"&nbsp;&nbsp;<font color='#cc0099'>"&city2&"</font>"&"<br>"
body=body&" Booking requirement:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop_ask_city2&"</font>"&"<br>"
body=body&"Check-in Date: "&"&nbsp;&nbsp;<font color='#cc0099'>"&in_date2&"</font>"&"<br>"
body=body&"Check-out Date: "&"&nbsp;&nbsp;<font color='#cc0099'>"&out_date2&"</font>"&"<br>"
body=body&"Grade:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop2_level&"</font>"&"<br>"
body=body&"No. of rooms:"&"&nbsp;&nbsp;Single:<font color='#cc0099'>"&city2_dr&"</font>room;"
body=body&"&nbsp;&nbsp;Double:<font color='#cc0099'>"&city2_sr&"</font>room;"
body=body&"&nbsp;&nbsp;Suite:<font color='#cc0099'>"&city2_tf&"</font>room<br><br>"
end if

if city3<>"" then
body=body&"City #3::"&"&nbsp;&nbsp;<font color='#cc0099'>"&city3&"</font>"&"<br>"
body=body&"Booking requirement:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop_ask_city3&"</font>"&"<br>"
body=body&"Check-in Date: "&"&nbsp;&nbsp;<font color='#cc0099'>"&in_date3&"</font>"&"<br>"
body=body&"Check-out Date: "&"&nbsp;&nbsp;<font color='#cc0099'>"&out_date3&"</font>"&"<br>"
body=body&"Grade:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop3_level&"</font>"&"<br>"
body=body&"No. of rooms:"&"&nbsp;&nbsp;Single:<font color='#cc0099'>"&city3_dr&"</font>room;"
body=body&"&nbsp;&nbsp;Double:<font color='#cc0099'>"&city3_sr&"</font>room;"
body=body&"&nbsp;&nbsp;Suite:<font color='#cc0099'>"&city3_tf&"</font>room<br><br>"
end if

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
 JMail.silent=true      '不将错误返回给操作系统
 JMail.Logging=True     '启动邮件日志
 JMail.ContentType = "text/html" 
 JMail.Charset="gb2312"    '邮件的文字编码为国标
 JMail.From=""&email&""    '发件人的E-MAIL地址,请修改成你的邮箱地址
 JMail.fromname=""&contact_name&""             '发件人的姓名
 JMail.mailserverusername=""&login_name&""    '发件人登录邮件服务器所需的用户名,请修改该项
 JMail.MailserverPassword=""&login_pass&""      '发件人登录邮件服务器所需的密码
 JMail.addRecipient ""&email&""     '收件人的E-MAIL地址,请修改该项
 JMail.Subject=""&contact_name&"Hotel Booking"&""      '邮件标题
 JMail.Body=""&body&""
 ifsend=JMail.send (""&server_smtp&"")    '执行发邮件,请修改成你邮箱的SMTP服务器
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
