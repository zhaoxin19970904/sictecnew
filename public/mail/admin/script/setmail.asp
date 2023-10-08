<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>设置邮箱地址</title>
</head>
<body>
<%
set rs=server.createobject("adodb.recordset")
sql="select id from e_mail where ifuse=true"
rs.open sql,conn,1,1
if not rs.eof then
%>
<script language="javascript">
  alert("你已设置了邮箱,如果要用新的请修改它。")
  location.href="changemail.asp"
</script>
<%
rs.close
response.end()
end if
setmail=request.form("set_ok")
if setmail="setmail" then
  email=request.form("email")
  login_name=request.form("login_name")
  login_pass=request.form("login_pass")
  server_smtp=request.form("server_smtp")
  sql="insert into e_mail(email,login_name,login_pass,server_smtp,ifuse) values('"&email&"','"&login_name&"','"&login_pass&"','"&server_smtp&"',true)"
  conn.execute(sql)
  conn.close
%>
<p>
  <script language="">
  alert("已成功设置了邮箱!")
  location.href="delmail.asp"
</script>
  <%
end if
%>
</p>
<p align="center"><strong><font size="+3">设 置 邮 箱</font></strong></p>
<form name="form1" method="post" action="" onsubmit="return check(this)">
  <table width="417" height="229" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="178" height="16">收邮件的邮箱:</td>
      <td width="219"><input name="email" type="text" id="email"> <font color="#FF0000">*</font></td>
      <td width="20">&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">该邮箱的用户名:</td>
      <td><input name="login_name" type="text" id="login_name"> <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">该邮箱的密码:</td>
      <td><input name="login_pass" type="password" id="login_pass"> <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">该邮箱的SMTP服务器:</td>
      <td><input name="server_smtp" type="text" id="server_smtp"> <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">&nbsp;</td>
      <td><input type="submit" name="Submit" value="确 定"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit2" value="重 写"></td>
      <td><input name="set_ok" type="hidden" id="set_ok" value="setmail"></td>
    </tr>
  </table>
</form>
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
