<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
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

company=request.form("company")

website=request.form("website")
adress=request.form("adress")
tel=request.form("tel")
fax=request.form("fax")
contactperson=request.form("contactperson")
position=request.form("position")
mobile=request.form("mobile")
e_mail=request.form("email")

no_employees=request.form("no_employees")
mianproducts=request.form("mianproducts")
annualproduction=request.form("annualproduction")
exporthistory=request.form("exporthistory")
area=request.form("area")
registered=request.form("registered")
certificate=request.form("certificate")
competitive=request.form("competitive")
notes=request.form("notes")

body="<div align='center'><font color='#ff9900' size=4>Become Our Suppliers</font></div>"&"<br>"
body=body&"Company Name:"&"&nbsp;&nbsp;<font color='#cc0099'>"&company&"</font>"&"<br>"

body=body&"Web site:"&"&nbsp;&nbsp;<font color='#cc0099'>"&website&"</font>"&"<br>"
body=body&"Adress:"&"&nbsp;&nbsp;<font color='#cc0099'>"&adress&"</font>"&"<br>"
body=body&"Telephone:"&"&nbsp;&nbsp;<font color='#cc0099'>"&tel&"</font>"&"<br>"
body=body&"Fax:"&"&nbsp;&nbsp;<font color='#cc0099'>"&fax&"</font>"&"<br>"
body=body&"Contact Person:"&"&nbsp;&nbsp;<font color='#cc0099'>"&contactperson&"</font>"&"<br>"
body=body&"Position:"&"&nbsp;&nbsp;<font color='#cc0099'>"&position&"</font>"&"<br>"
body=body&"Mobile:"&"&nbsp;&nbsp;<font color='#cc0099'>"&mobile&"</font>"&"<br>"
body=body&"E-mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br>"

body=body&"Number of employees:"&"&nbsp;&nbsp;<font color='#cc0099'>"&no_employees&"</font>"&"<br>"
body=body&"Mian products:"&"&nbsp;&nbsp;<font color='#cc0099'>"&mianproducts&"</font>"&"<br>"
body=body&"Annual production:"&"&nbsp;&nbsp;<font color='#cc0099'>"&annualproduction&"</font>"&"<br>"
body=body&"Export history:"&"&nbsp;&nbsp;<font color='#cc0099'>"&exporthistory&"</font>"&"<br>"
body=body&"Major Export Area:"&"&nbsp;&nbsp;<font color='#cc0099'>"&area&"</font>"&"<br><br>"
body=body&"ISO registered:"&"&nbsp;&nbsp;<font color='#cc0099'>"&registered&"</font>"&"<br><br>"
body=body&"Certificates/Approvals:"&"&nbsp;&nbsp;<font color='#cc0099'>"&certificate&"</font>"&"<br><br>"
body=body&"Competitive Strengh:"&"&nbsp;&nbsp;<font color='#cc0099'>"&competitive&"</font>"&"<br><br>"
body=body&"Notes/Comments:"&"&nbsp;&nbsp;<font color='#cc0099'>"&notes&"</font>"&"<br><br>"

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
 JMail.Subject=""&"Become our Suppliers from  "&company&""      
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
