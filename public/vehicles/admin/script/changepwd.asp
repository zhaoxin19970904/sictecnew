<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%response.buffer=true%>
<%
if session("admin_name")="" then response.end
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改管理密码</title>
<script language="jscript">
<!--
   function check(form1)
   {
     var NotNull
	 NotNull=true
	   if(form1.username.value=="")
	     {
		   window.alert("Please input New User Name！")
		   NotNull=false
		 }
		if(form1.password.value=="")
		 {
		   window.alert("Please input New password！")
		   NotNull=false
		 }
		 if(form1.oldpwd.value=="")
		 {
		   window.alert("Please input old password ！")
		   NotNull=false
		 }
		 if(form1.confirmpwd.value=="")
		 {
		   window.alert("Please input affirm password！")
		   NotNull=false
		 }
		 if(form1.password.value!=form1.confirmpwd.value)
		 {
		   window.alert("New password different from affirm password！")
		   NotNull=false
		 }
		 return NotNull
   }
//-->
</script>
<style type="text/css">
<!--

.dateformat {
	font-size: 13px;
	color: #383805;
}
	 a:link {text-decoration:none; color:#383805;font-size=12px}
	 a:visited {text-decoration:none; color:#383805;font-size=12px}
     a:active {text-decoration:underline; color:#383805}
     a:hover {text-decoration:underline; color:#FF0033}
-->
</style>
</head>

<body topmargin="50">
<%
if request.form("changepwd")<>"" then
  dim rs,sql
  set rs=server.CreateObject("adodb.recordset")
  sql="select admin_name,admin_pass from admin where admin_name='"&session("admin_name")&"'"
  rs.open sql,conn,1,3
  if trim(request.form("oldpwd"))<>rs("admin_pass") then
   response.write "Old passowrd is wrong!<br><br>"
   response.write "<a href='#' onclick='javascript:history.back()'>Back</a>"
  else
   rs("admin_name")=trim(request.form("username"))
   rs("admin_pass")=trim(request.form("password"))
   rs.update
   rs.close
   set rs=nothing
   conn.close
   set conn=nothing
   response.write "Successfully！<br><br>"
   response.write "<a href='../index.htm' target='_top'>Confirm</a>"
  end if
else
%>
<form name="form1" method="post" action="" onsubmit="return check(this)">
  <table width="515" height="168" border="0" align="center" cellpadding="0" cellspacing="0" background="../images/bg_dir.jpg" class="dateformat">
    <tr> 
      <td colspan="2"> 
        <div align="center">
		<font size="6" face="仿宋_GB2312">Amend Manage PWD</font>
		<img src="../../images/line_01.jpg" width="338" height="20"></div>
		</td>
    </tr>
    <tr> 
      <td width="230" height="32"> 
	  <div align="right">New User Name: </div>
	  </td>
      <td width="285"> 
        <input name="username" type="text" size="20" style="ime-mode:disabled" onkeydown="if(event.keyCode==13)event.keyCode=9">
      </td>
    </tr>
    <tr>
      <td height="35"><div align="right">New Password：</div></td>
      <td><input name="password" type="password" id="password" size="20" style="ime-mode:disabled" onkeydown="if(event.keyCode==13)event.keyCode=9"></td>
    </tr>
    <tr> 
      <td height="35"><div align="right">Affirm Password：</div></td>
      <td><input name="confirmpwd" type="password" id="confirmpwd" size="20" style="ime-mode:disabled" onkeydown="if(event.keyCode==13)event.keyCode=9"></td>
    </tr>
    <tr> 
      <td height="35"> <div align="right">Old Password：</div></td>
      <td><input name="oldpwd" type="password" id="oldpwd" style="ime-mode:disabled" onkeydown="if(event.keyCode==13)event.Submit.onkeydown" size="20"> 
        <input name="changepwd" type="hidden" id="changepwd" value="changepwd15"></td>
    </tr>
    <tr> 
      <td height="39">&nbsp;</td>
      <td><input type="submit" name="submit" value="Submit"> &nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit2" value="Reset"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<%end if%>
</body>
</html>
