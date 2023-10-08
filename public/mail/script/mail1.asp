<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>发送邮件</title>
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
e_mail=request.form("e_mail")
if IsValidEmail(e_mail)=false then
  %>
 <script language="javascript">
   alert("邮件的格式错误,正确的格式如:\nabc123@126.com")
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
city4=request.form("city4")
city5=request.form("city5")

if city1<>"" then
  shop_ask_city1=request.form("shop_ask_city1")
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
  shop3_level=request.Form("shop3_level")
  city3_dr=request.Form("city3_dr")
  city3_sr=request.Form("city3_sr")
  city3_tf=request.Form("city3_tf")
  if city3_dr="" then city3_dr="0"
  if city3_sr="" then city3_sr="0"
  if city3_tf="" then city3_tf="0"
end if

if city4<>"" then
  shop_ask_city4=request.form("shop_ask_city4")
  shop4_level=request.Form("shop4_level")
  city4_dr=request.Form("city4_dr")
  city4_sr=request.Form("city4_sr")
  city4_tf=request.Form("city4_tf")
  if city4_dr="" then city4_dr="0"
  if city4_sr="" then city4_sr="0"
  if city4_tf="" then city4_tf="0"
end if

if city5<>"" then
  shop_ask_city5=request.form("shop_ask_city5")
  shop5_level=request.Form("shop5_level")
  city5_dr=request.Form("city5_dr")
  city5_sr=request.Form("city5_sr")
  city5_tf=request.Form("city5_tf")
  if city5_dr="" then city5_dr="0"
  if city5_sr="" then city5_sr="0"
  if city5_tf="" then city5_tf="0"
end if


body="<div align='center'><font color='#ff9900' size=4>预订酒店</font></div>"&"<br>"
body=body&"顾客的姓名:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_name&"</font>"&"<br>"
body=body&"顾客的单位:"&"&nbsp;&nbsp;<font color='#cc0099'>"&company&"</font>"&"<br>"
body=body&"顾客的职位:"&"&nbsp;&nbsp;<font color='#cc0099'>"&place&"</font>"&"<br>"
body=body&"顾客的联系电话:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_tel&"</font>"&"<br>"
body=body&"顾客的E-mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br><br>"

if city1<>"" then
body=body&"顾客预定的城市一:"&"&nbsp;&nbsp;<font color='#cc0099'>"&city1&"</font>"&"<br>"
body=body&"顾客对酒店的要求:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop_ask_city1&"</font>"&"<br>"
body=body&"酒店的级别:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop1_level&"</font>"&"<br>"
body=body&"酒店的间数:"&"&nbsp;&nbsp;单人房:<font color='#cc0099'>"&city1_dr&"</font>间;"
body=body&"&nbsp;&nbsp;双人房:<font color='#cc0099'>"&city1_sr&"</font>间;"
body=body&"&nbsp;&nbsp;套房:<font color='#cc0099'>"&city1_tf&"</font>间<br><br>"
end if

if city2<>"" then
body=body&"顾客预定的城市二:"&"&nbsp;&nbsp;<font color='#cc0099'>"&city2&"</font>"&"<br>"
body=body&"顾客对酒店的要求:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop_ask_city2&"</font>"&"<br>"
body=body&"酒店的级别:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop2_level&"</font>"&"<br>"
body=body&"酒店的间数:"&"&nbsp;&nbsp;单人房:<font color='#cc0099'>"&city2_dr&"</font>间;"
body=body&"&nbsp;&nbsp;双人房:<font color='#cc0099'>"&city2_sr&"</font>间;"
body=body&"&nbsp;&nbsp;套房:<font color='#cc0099'>"&city2_tf&"</font>间<br><br>"
end if

if city3<>"" then
body=body&"顾客预定的城市三:"&"&nbsp;&nbsp;<font color='#cc0099'>"&city3&"</font>"&"<br>"
body=body&"顾客对酒店的要求:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop_ask_city3&"</font>"&"<br>"
body=body&"酒店的级别:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop3_level&"</font>"&"<br>"
body=body&"酒店的间数:"&"&nbsp;&nbsp;单人房:<font color='#cc0099'>"&city3_dr&"</font>间;"
body=body&"&nbsp;&nbsp;双人房:<font color='#cc0099'>"&city3_sr&"</font>间;"
body=body&"&nbsp;&nbsp;套房:<font color='#cc0099'>"&city3_tf&"</font>间<br><br>"
end if

if city4<>"" then
body=body&"顾客预定的城市四:"&"&nbsp;&nbsp;<font color='#cc0099'>"&city4&"</font>"&"<br>"
body=body&"顾客对酒店的要求:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop_ask_city4&"</font>"&"<br>"
body=body&"酒店的级别:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop4_level&"</font>"&"<br>"
body=body&"酒店的间数:"&"&nbsp;&nbsp;单人房:<font color='#cc0099'>"&city4_dr&"</font>间;"
body=body&"&nbsp;&nbsp;双人房:<font color='#cc0099'>"&city4_sr&"</font>间;"
body=body&"&nbsp;&nbsp;套房:<font color='#cc0099'>"&city4_tf&"</font>间<br><br>"
end if

if city5<>"" then
body=body&"顾客预定的城市五:"&"&nbsp;&nbsp;<font color='#cc0099'>"&city5&"</font>"&"<br>"
body=body&"顾客对酒店的要求:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop_ask_city5&"</font>"&"<br>"
body=body&"酒店的级别:"&"&nbsp;&nbsp;<font color='#cc0099'>"&shop5_level&"</font>"&"<br>"
body=body&"酒店的间数:"&"&nbsp;&nbsp;单人房:<font color='#cc0099'>"&city5_dr&"</font>间;"
body=body&"&nbsp;&nbsp;双人房:<font color='#cc0099'>"&city5_sr&"</font>间;"
body=body&"&nbsp;&nbsp;套房:<font color='#cc0099'>"&city5_tf&"</font>间<br><br>"
end if


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
 JMail.silent=true      '不将错误返回给操作系统
 JMail.Logging=True     '启动邮件日志
 JMail.ContentType = "text/html" 
 JMail.Charset="gb2312"    '邮件的文字编码为国标
 JMail.From=""&email&""    '发件人的E-MAIL地址,请修改成你的邮箱地址
 JMail.fromname=""&contact_name&""             '发件人的姓名
 JMail.mailserverusername=""&login_name&""    '发件人登录邮件服务器所需的用户名,请修改该项
 JMail.MailserverPassword=""&login_pass&""      '发件人登录邮件服务器所需的密码
 JMail.addRecipient ""&email&""     '收件人的E-MAIL地址,请修改该项
 JMail.Subject=""&company&""     '邮件标题
 JMail.Body=""&body&""
 ifsend=JMail.send (""&server_smtp&"")    '执行发邮件,请修改成你邮箱的SMTP服务器
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
   location.href="../index1.asp"
</script>
<%end if%>
</body>
</html>
