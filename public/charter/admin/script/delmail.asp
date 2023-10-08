<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>
<%
mailid=request.form("mailid")
if mailid<>"" then
sql="delete * from e_mail where id="&mailid&""
conn.execute(sql)
%>
<script language="javascript">
  alert(""邮箱已成功删除！)
  location.href="setmail.asp"
</script>
<%
end if
sql="select * from e_mail where ifuse=true"
set rs=conn.execute(sql)
if not rs.eof then
%>
<p align="center"><strong><font size="+3">删 除 邮 箱</font></strong></p>
<form name="form1" method="post" action="" onsubmit="return check(this)">
  <table width="500" height="229" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="178" height="16">收邮件的邮箱:</td>
      <td width="297">
	 <%=rs("email")%>
	 </td>
      <td width="25">&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">该邮箱的用户名:</td>
      <td><%=rs("login_name")%></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">该邮箱的密码:</td>
      <td>********</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">该邮箱的SMTP服务器:</td>
      <td><%=rs("server_smtp")%></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="16">&nbsp;</td>
      <td><input type="submit" name="Submit" value="确 定 删 除 ">
      </td>
      <td> 
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
  alert("没有要删除的邮箱，请先设置!")
  location.href="setmail.asp"
</script>
<%
end if
%>
</body>
</html>
