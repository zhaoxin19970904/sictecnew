<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%response.buffer=true%>
<%
if session("admin_name")="" then response.end
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸Ĺ�������</title>
<script language="jscript">
<!--
   function check(form1)
   {
     var NotNull
	 NotNull=true
	   if(form1.username.value=="")
	     {
		   window.alert("�û�������Ϊ�գ�")
		   NotNull=false
		 }
		if(form1.password.value=="")
		 {
		   window.alert("�����벻��Ϊ�գ�")
		   NotNull=false
		 }
		 if(form1.oldpwd.value=="")
		 {
		   window.alert("�����벻��Ϊ�գ�")
		   NotNull=false
		 }
		 if(form1.confirmpwd.value=="")
		 {
		   window.alert("����ȷ�ϲ���Ϊ�գ�")
		   NotNull=false
		 }
		 if(form1.password.value!=form1.confirmpwd.value)
		 {
		   window.alert("������������벻��ͬ��")
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
   response.write "�����벻��ȷ"
   response.write "<a href='#' onclick='javascript:history.back()'>�� ��</a>"
  else
   rs("admin_name")=trim(request.form("username"))
   rs("admin_pass")=trim(request.form("password"))
   rs.update
   rs.close
   set rs=nothing
   conn.close
   set conn=nothing
   response.write "�޸Ĺ�������ɹ���"
   response.write "<a href='../index.htm' target='_top'>ȷ ��</a>"
  end if
else
%>
<form name="form1" method="post" action="" onsubmit="return check(this)">
  <table width="400" height="168" border="0" align="center" cellpadding="0" cellspacing="0" background="../images/bg_dir.jpg" class="dateformat">
    <tr> 
      <td colspan="2"> 
        <div align="center"><font size="6" face="����_GB2312">�޸Ĺ�������<img src="../../images/line_01.jpg" width="338" height="20"></font></div></td>
    </tr>
    <tr> 
      <td width="182" height="32"> <div align="right">�û�����</div></td>
      <td width="218"> 
        <input name="username" type="text" size="20" style="ime-mode:disabled" onkeydown="if(event.keyCode==13)event.keyCode=9">
      </td>
    </tr>
    <tr>
      <td height="35"><div align="right">�����룺</div></td>
      <td><input name="password" type="password" id="password" size="20" style="ime-mode:disabled" onkeydown="if(event.keyCode==13)event.keyCode=9"></td>
    </tr>
    <tr> 
      <td height="35"><div align="right">����ȷ�ϣ�</div></td>
      <td><input name="confirmpwd" type="password" id="confirmpwd" size="20" style="ime-mode:disabled" onkeydown="if(event.keyCode==13)event.keyCode=9"></td>
    </tr>
    <tr> 
      <td height="35"> <div align="right">�����룺</div></td>
      <td><input name="oldpwd" type="password" id="oldpwd" style="ime-mode:disabled" onkeydown="if(event.keyCode==13)event.Submit.onkeydown" size="20"> 
        <input name="changepwd" type="hidden" id="changepwd" value="changepwd15"></td>
    </tr>
    <tr> 
      <td height="39">&nbsp;</td>
      <td><input type="submit" name="submit" value="ȷ��"> &nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit2" value="��д"></td>
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
