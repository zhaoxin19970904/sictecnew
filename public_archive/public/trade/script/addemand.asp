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
p_type=request.Form("type")
pro_name=request.Form("pro_name")
price=request.Form("price")
d_time=request.Form("time")
notes=request.Form("notes")
contact_name=request.Form("name")
tel=request.Form("tel")
set rs=server.CreateObject("adodb.recordset")
sql="select * from demand"
rs.open sql,conn,1,3
rs.addnew
rs("memberid")=session("username")
rs("type")=p_type
rs("pro_name")=pro_name
rs("price")=price
rs("time")=d_time
rs("demand_time")=now()
rs("notes")=notes
rs("name")=contact_name
rs("tel")=tel
rs("email")=email
rs("feedback")=false
rs("demand_ip")=request.ServerVariables("REMOTE_ADDR")
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script language="JavaScript">
   alert("Successfully!")
   location.href="../demand.asp"
</script>
</body>
</html>
