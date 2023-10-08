<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Member register</title>
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

'member_id=request.form("member_id")
member_id=trim(session("username"))
if member_id="" then 
%>

<script language="JavaScript">
   alert("Time out,pls login first!")
    location.href="../index.asp"
      </script>

<%
response.end()
end if



e_mail=request.form("email")
psd_f=trim(request.form("psd_f"))


sql="select psd from member where member_uid='"&member_id&"'"
set rs=conn.execute(sql)

if psd_f<>"" then

psd=trim(request.form("psd"))
re_psd=trim(request.form("re_psd"))

if rs("psd")=psd_f and psd=re_psd then
m=1
else 
 %>

<script language="JavaScript">
   alert("The former password you entered or the re-entered new password is incorrect,pls check!")
   history.back()
   </script>


<%
response.end()
end if
 %>



<%

end if

sql="select id from member where member_uid='"&member_id&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language="JavaScript">
   alert("No such user,pls try later!")
   history.back()
   </script>
<%
response.end()
end if
rs.close
 %>

<%

if IsValidEmail(e_mail)=false then
  %>
 <script language="javascript">
   alert("Wrong Email address,pls check!")
   history.back()
 </script>
 

<%
response.end()
%>
 
 <% 

end if %>


<%

firstname=request.form("firstname")
lastname=request.form("lastname")
'sex=request.form("sex")
'female=request.form("female")
jobtitle=request.form("jobtitle")
tel=request.form("tel")
'e_mail=request.form("email")
country=request.form("country")
city=request.form("city")
address=request.form("address")
post=request.form("post")

company=request.form("company")
area=request.form("area")
main_products=request.form("main_products")
biznature=request.form("biznature")
com_tel=request.form("com_tel")
fax=request.form("fax")
productbuy=request.form("productbuy")
productsell=request.form("productsell")
homepage=request.form("homepage")

%>


 
 <%
  
  set rs=server.createobject("adodb.recordset")
  sql="select * from member where member_uid='"&member_id&"'"
  rs.open sql,conn,1,3
  'rs.addnew
  'rs("member_uid")=member_id
  if m=1 then 
  rs("psd")=psd
  end if
  'rs("sex")=sex
  rs("firstname")=firstname
  rs("lastname")=lastname
  rs("jobtitle")=jobtitle
  rs("email")=e_mail
  rs("tel")=tel
  rs("country")=country
  rs("city")=city
  rs("address")=address
  rs("post")=post
  rs("company")=company
  rs("main_biz_area")=area
  rs("main_products")=main_products
  rs("biznature")=biznature
  rs("com_tel")=com_tel
  rs("fax")=fax
  rs("productbuy")=productbuy
  rs("productsell")=productsell
  rs("homepage")=homepage
  'rs("reg_time")=now()
  'rs("reg_ip")=request.ServerVariables("REMOTE_ADDR")
  rs.update
  'session("username")=member_id
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  
%>

<script language="JavaScript">
  alert("Update Successfully!")
</script>



<%

body="<div align='center'><font color='#ff9900' size=4>[System Notification] A member has just updated his profile on trade-universal.com</font></div>"&"<br>"

body=body&"Member ID:"&"&nbsp;&nbsp;<font color='#cc0099'>"&member_id&"</font>"&"<br>"

body=body&"Name: "&"&nbsp;&nbsp;<font color='#cc0099'>"&firstname&"</font>"&"<br>"
body=body&"Name: "&"&nbsp;&nbsp;<font color='#cc0099'>"&lastname&"</font>"&"<br>"
'body=body&"Sex:"&"&nbsp;&nbsp;<font color='#cc0099'>"&sex&"</font>"&"<br>"
'body=body&"Sex:"&"&nbsp;&nbsp;<font color='#cc0099'>"&female&"</font>"&"<br>"
body=body&"Job title:"&"&nbsp;&nbsp;<font color='#cc0099'>"&jobtitle&"</font>"&"<br>"
body=body&"Contact tel:"&"&nbsp;&nbsp;<font color='#cc0099'>"&tel&"</font>"&"<br>"
body=body&"E-Mail:"&"&nbsp;&nbsp;<font color='#cc0099'>"&e_mail&"</font>"&"<br>"
body=body&"Country/region:"&"&nbsp;&nbsp;<font color='#cc0099'>"&country&"</font>"&"<br>"
body=body&"City:"&"&nbsp;&nbsp;<font color='#cc0099'>"&city&"</font>"&"<br>"
body=body&"Address:"&"&nbsp;&nbsp;<font color='#cc0099'>"&address&"</font>"&"<br>"
body=body&"Postal code:"&"&nbsp;&nbsp;<font color='#cc0099'>"&post&"</font>"&"<br>"

body=body&"Company name:"&"&nbsp;&nbsp;<font color='#cc0099'>"&company&"</font>"&"<br>"
body=body&"Major business area:"&"&nbsp;&nbsp;<font color='#cc0099'>"&area&"</font>"&"<br>"
body=body&"Main products/services:"&"&nbsp;&nbsp;<font color='#cc0099'>"&main_products&"</font>"&"<br>"
body=body&"Business nature:"&"&nbsp;&nbsp;<font color='#cc0099'>"&biznature&"</font>"&"<br>"
body=body&"Tel:"&"&nbsp;&nbsp;<font color='#cc0099'>"&com_tel&"</font>"&"<br>"
body=body&"Fax:"&"&nbsp;&nbsp;<font color='#cc0099'>"&fax&"</font>"&"<br>"
body=body&"Products to purchase:"&"&nbsp;&nbsp;<font color='#cc0099'>"&productbuy&"</font>"&"<br>"
body=body&"Products to sell:"&"&nbsp;&nbsp;<font color='#cc0099'>"&productsell&"</font>"&"<br>"
body=body&"Homepage:"&"&nbsp;&nbsp;<font color='#cc0099'>"&homepage&"</font>"&"<br>"


%>

<%
		
	db1="../../../databases/mail.mdb"
	
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db1)
	
	'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(db)
	conn.Open connstr


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
 JMail.Subject=""&"Quotation from  "&company&""      
 JMail.Body=""&body&""
 ifsend=JMail.send (""&server_smtp&"")    
 JMail.close
 set JMail=nothing
 if ifsend=true then
    send="Registration Successful ,Thanks!"
  else
	send="Time out. If this occurs constantly,pls contact us by trade@sictec.com, thanks!"
 end if
%>

<%
  conn.close
  set conn=nothing
  %>


<script language="JavaScript">
  location.href="../index.asp"
</script>

</body>
</html>
