<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改邮箱地址</title>
</head>
<body>
<%
changemail=request.form("changemail")
mailid=request.form("mailid")
if changemail="changemail" then
  email=request.form("email")
  login_name=request.form("login_name")
  login_pass=request.form("login_pass")
  server_smtp=request.form("server_smtp")
  sql="select * from e_mail where id="&mailid&""
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,3
  rs("email")=email
  rs("login_name")=login_name
  rs("login_pass")=login_pass
  rs("server_smtp")=server_smtp
  rs.update
  rs.close
  set rs=nothing
%>
  <script language="">
  alert("已成功修改了邮箱!")
  location.href="changemail.asp"
</script>
  <%
end if
sql="select * from e_mail where ifuse=true"
set rs=conn.execute(sql)
if not rs.eof then
%>
<p align="center"><strong><font size="+3">修 改 邮 箱</font></strong></p>
<form name="form1" method="post" action="" onsubmit="return check(this)">
  <table width="500" height="229" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="178" height="16">收邮件的邮箱:</td>
      <td width="297">
	  <input name="email" type="text" id="email" value="<%=rs("email")%>"> <font color="#FF0000">*</font></td>
      <td width="25">&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">该邮箱的用户名:</td>
      <td><input name="login_name" type="text" id="login_name" value="<%=rs("login_name")%>"> <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">该邮箱的密码:</td>
      <td><input name="login_pass" type="password" id="login_pass" value="<%=rs("login_pass")%>"> <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">该邮箱的SMTP服务器:</td>
      <td><input name="server_smtp" type="text" id="server_smtp" value="<%=rs("server_smtp")%>"> <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">&nbsp;</td>
      <td><input type="submit" name="Submit" value="确 定 修 改"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit2" value="重 写"></td>
      <td><input name="changemail" type="hidden" id="changemail" value="changemail">
        <input name="mailid" type="hidden" id="mailid" value="<%=rs("id")%>"></td>
    </tr>
  </table>
  <%
  rs.close
  conn.close
  %>
</form>
<%
else
%>
 <script language="">
  alert("没有要修改的邮箱，请先设置!")
  location.href="setmail.asp"
</script>
<%
end if
%>
<script language="JScript">
<!--
function check(form1)
{
 
if (form1.email.value=="")
 {
    alert("收邮件的邮箱!")
	form1.email.focus()
    return false
  }
 
  if (form1.login_name.value=="")
 {
    alert("该邮箱的用户名!")
	form1.login_name.focus()
    return false
  }
 
  
  if (form1.login_pass.value=="")
 {
    alert("该邮箱的密码!")
	form1.login_pass.focus()
    return false
  }
 
  
  if (form1.server_smtp.value=="")
 {
    alert("该邮箱的SMTP服务器!")
	form1.server_smtp.focus()
    return false
  }
  
}
//-->
</script>
</body>
</html>
