<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>tire clock order</title>
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
company=request.form("company")
tel=request.form("tel")
fax=request.form("fax")
e_mail=request.form("email")
address=request.form("address")

TCT_455=request.form("TCT_455")
TCT_455pcs=request.form("TCT_455pcs")

TCW_400A=request.form("TCW_400A")
TCW_400Apcs=request.form("TCW_400Apcs")

TCW_901=request.form("TCW_901")
TCW_901pcs=request.form("TCW_901pcs")

TCW_405=request.form("TCW_405")
TCW_405pcs=request.form("TCW_405pcs")

TCW_435=request.form("TCW_435")
TCW_435pcs=request.form("TCW_435pcs")

TCW_510=request.form("TCW_510")
TCW_510pcs=request.form("TCW_510pcs")

TCW_520=request.form("TCW_520")
TCW_520pcs=request.form("TCW_520pcs")

TCK_470=request.form("TCK_470")
TCK_470pcs=request.form("TCK_470pcs")

TCT_530=request.form("TCT_530")
TCT_530pcs=request.form("TCT_530pcs")

TCT_540=request.form("TCT_540")
TCT_540pcs=request.form("TCT_540pcs")

TCT_550=request.form("TCT_550")
TCT_550pcs=request.form("TCT_550pcs")

TCT_560=request.form("TCT_560")
TCT_560pcs=request.form("TCT_560pcs")

TCT_570=request.form("TCT_570")
TCT_570pcs=request.form("TCT_570pcs")

TCT_580=request.form("TCT_580")
TCT_580pcs=request.form("TCT_580pcs")

TCT_590=request.form("TCT_590")
TCT_590pcs=request.form("TCT_590pcs")

TCT_600=request.form("TCT_600")
TCT_600pcs=request.form("TCT_600pcs")

notes=request.form("notes")

body="<div align='center'><font color='#ff9900' size=4>Tire Clock Order </font></div>"&"<br>"
body=body&"Name:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contact_name&"</font>"&"<br>"
body=body&"Company:"&"&nbsp;&nbsp;<font color='#cc0099'>"&company&"</font>"&"<br>"
body=body&"Tel:"&"&nbsp;&nbsp;<font color='#cc0099'>"&tel&"</font>"&"<br>"
body=body&"Fax:"&"&nbsp;&nbsp;<font color='#cc0099'>"&fax&"</font>"&"<br>"
body=body&"E-mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br><br>"
body=body&"Address:"&"&nbsp;&nbsp;<font color='#cc0099'>"&address&"</font>"&"<br>"


body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_455&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_455pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_400A&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_400Apcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_901&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_901pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_405&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_405pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_435&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_435pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_510&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_510pcs&"</font>"&"<br>"


body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_520&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCW_520pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCK_470&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCK_470pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_530&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_530pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_540&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_540pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_550&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_550pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_560&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_560pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_570&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_570pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_580&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_580pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_590&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_590pcs&"</font>"&"<br>"

body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_600&"</font>"&"<br>"
body=body&"Products Name And quantity:"&"&nbsp;&nbsp;<font color='#cc0099'>"&TCT_600pcs&"</font>"&"<br>"

body=body&"Notes &amp; Comments:"&"&nbsp;&nbsp;<font color='#cc0099'>"&notes&"</font>"&"<br><br>"


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
 JMail.Subject=""&"Ordering Tire Clock from  "&contact_name&""      
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
