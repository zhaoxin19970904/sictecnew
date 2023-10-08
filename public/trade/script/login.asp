<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Login</title>
</head>
<body>
<%
if session("username")<>"" then
  session.Abandon()
  %>
<script language="JavaScript">
alert("You have login,pls exit first!")
location.href="../index.asp"
</script>
  
  <%
    response.End()
else
  username=trim(request.form("username"))
  userpass=trim(request.form("userpass"))
  sql="select psd from member where member_uid='"&username&"'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then 
%>
<script language="JavaScript">
alert("No such member.Please register first!")
location.href="../register.asp"
</script>
<%
    'response.write "Sorry,you've not registered yet."
    else
     if trim(rs("psd"))=userpass then
	     session("username")=username
         response.redirect("../index.asp")
       else
%>
<script language="JavaScript">
alert("Wrong username or password.Pls try again!")
history.back()
</script>
<%
         'response.write "Wrong username or or password.Pls try again!"
     end if
  end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>
</body>
</html>
