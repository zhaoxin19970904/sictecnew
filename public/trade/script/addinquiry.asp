<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Add Demand</title>
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
%>
<%
email=trim(request.Form("email"))
if email<>"" then
    if IsValidEmail(email)=false then
%>
<script language="JavaScript">
   alert("Your E-mail is Wrong")
   location.href="../demand.asp"
</script>
<%
  response.End()
   end if
end if
i_date=request.Form("date")
subject=request.Form("subject")
productid=request.Form("productid")
price=request.Form("price")
packing=request.Form("packing")
minimum=request.Form("minimum")
product=request.Form("product")
terms=request.Form("terms")
discount=request.Form("discount")
packingdetails=request.Form("packingdetails")
delivery=request.Form("delivery")
toknow=price&","&packing&","&minimum&","&product&","&terms&","&discount&","&packingdetails&","&delivery


order=request.Form("order")
notes=request.Form("notes")
contactperson=request.Form("contactperson")
jobtitle=request.Form("jobtitle")
tel=request.Form("tel")
set rs=server.CreateObject("adodb.recordset")
sql="select * from inquiry"
rs.open sql,conn,1,3
rs.addnew
rs("memberid")=session("username")
rs("date")=i_date
rs("subject")=subject
rs("productid")=productid
rs("toknow")=toknow
rs("order")=order
rs("notes")=notes
rs("contactperson")=contactperson
rs("jobtitle")=jobtitle
rs("tel")=tel
rs("email")=email
rs("feedback")=false
rs("inquiry_ip")=request.ServerVariables("REMOTE_ADDR")
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script language="JavaScript">
   alert("Successfully!")
   location.href="../inquiry.asp"
</script>
</body>
</html>
