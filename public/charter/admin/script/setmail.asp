<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���������ַ</title>
</head>
<body>
<%
set rs=server.createobject("adodb.recordset")
sql="select id from e_mail where ifuse=true"
rs.open sql,conn,1,1
if not rs.eof then
%>
<script language="javascript">
  alert("��������������,���Ҫ���µ����޸�����")
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
  alert("�ѳɹ�����������!")
  location.href="delmail.asp"
</script>
  <%
end if
%>
</p>
<p align="center"><strong><font size="+3">�� �� �� ��</font></strong></p>
<form name="form1" method="post" action="" onsubmit="return check(this)">
  <table width="417" height="229" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="178" height="16">���ʼ�������:</td>
      <td width="219"><input name="email" type="text" id="email"> <font color="#FF0000">*</font></td>
      <td width="20">&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">��������û���:</td>
      <td><input name="login_name" type="text" id="login_name"> <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">�����������:</td>
      <td><input name="login_pass" type="password" id="login_pass"> <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">�������SMTP������:</td>
      <td><input name="server_smtp" type="text" id="server_smtp"> <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">&nbsp;</td>
      <td><input type="submit" name="Submit" value="ȷ ��"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit2" value="�� д"></td>
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
    alert("���ʼ�������!")
	form1.email.focus()
    return false
  }
 
  if (form1.login_name.value=="")
 {
    alert("��������û���!")
	form1.login_name.focus()
    return false
  }
 
  
  if (form1.login_pass.value=="")
 {
    alert("�����������!")
	form1.login_pass.focus()
    return false
  }
 
  
  if (form1.server_smtp.value=="")
 {
    alert("�������SMTP������!")
	form1.server_smtp.focus()
    return false
  }
  
}
//-->
</script>
</body>
</html>
