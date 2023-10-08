<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Add Register Member</title>
</head>
<body>
<%
function IsValidEmail(user_mail)
dim names, name, i, c
IsValidEmail = true
names = Split(user_mail, "@")
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


email=trim(request("email"))
member_uid=trim(request.form("member_uid"))
sql="select id from member where member_uid='"&member_uid&"'"
set rs=conn.execute(sql)
if not rs.eof then
%>
<script language="JavaScript">
   alert("This member name have registered!")
   location.href="../register.asp"
</script>
<%
end if
rs.close
if email<>"" then
   if isvalidemail(email)=false then
%>
<script language="JavaScript">
   alert("Your E-mail is Wrong")
   location.href="../register.asp"
</script>
<%
     else
   end if
end if
 
  psd=trim(request.form("psd"))
  sex=request.form("sex")
  firstname=request.form("firstname")
  lastname=request.form("lastname")
  jobtitle=request.form("jobtitle")
  'email=request.form("email")
  tel=request.form("tel")
  country=request.form("country")
  city=request.form("city")
  address=request.form("address")
  post=request.form("post")
  company=request.form("company")
  main_biz_area=request.form("bigclass")
  main_products=request.form("main_products")
  biznature=request.form("biznature")
  com_tel=request.form("com_tel")
  fax=request.form("fax")
  productbuy=request.form("productbuy")
  productsell=request.form("productsell")
  homepage=request.form("homepage")
  set rs=server.createobject("adodb.recordset")
  sql="select * from member"
  rs.open sql,conn,1,3
  rs.addnew
  rs("member_uid")=member_uid
  rs("psd")=psd
  rs("sex")=sex
  rs("firstname")=firstname
  rs("lastname")=lastname
  rs("jobtitle")=jobtitle
  rs("email")=email
  rs("tel")=tel
  rs("country")=country
  rs("city")=city
  rs("address")=address
  rs("post")=post
  rs("company")=company
  rs("main_biz_area")=main_biz_area
  rs("main_products")=main_products
  rs("biznature")=biznature
  rs("com_tel")=com_tel
  rs("fax")=fax
  rs("productbuy")=productbuy
  rs("productsell")=productsell
  rs("homepage")=homepage
  rs("reg_time")=now()
  rs("reg_ip")=request.ServerVariables("REMOTE_ADDR")
  rs.update
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
%>
<script language="JavaScript">
  alert("Successfully Register!")
  location.href="../signin.asp"
</script>
</body>
</html>
