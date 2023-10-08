<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>包车服务</title>
<style type="text/css">
.text{
  font-size:12px;
}
</style>
</head>
<body>
<form name="form1" method="post" action="script/mail.asp" onsubmit="return check(this)">
  <table width="768" height="433" border="0" cellpadding="0" cellspacing="0" class="text">
    <tr bgcolor="#cccccc"> 
      <td colspan="2"> <font color="#0033CC"><strong>您 的 信 息</strong></font></td>
    </tr>
    <tr> 
      <td width="169"><div align="center">您的姓名:</div></td>
      <td width="599"><input name="book_name" type="text" id="book_name"> <font color="#CC0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">您的单位:</div></td>
      <td><input name="book_unit" type="text" id="book_unit">
        <font color="#CC0000">*</font> </td>
    </tr>
    <tr> 
      <td><div align="center">您的职位:</div></td>
      <td><input name="buyer_place" type="text" id="buyer_place"></td>
    </tr>
    <tr> 
      <td><div align="center">联系电话:</div></td>
      <td><input name="contact_tel" type="text" id="contact_tel"> <font color="#CC0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">E-mail:</div></td>
      <td><input name="email" type="text" id="email"> <font color="#CC0000">*</font></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td height="23" colspan="2"> <font color="#0033CC"><strong>包 车</strong></font></td>
    </tr>
    <tr> 
      <td><div align="center">包车日期:</div></td>
      <td><input name="book_date" type="text" id="book_date">
        <font color="#CC0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">包车路线:</div></td>
      <td><input name="book_way" type="text" id="book_way">
        <font color="#CC0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">包车车辆:</div></td>
      <td><input name="car7" type="checkbox" id="car7" value="7座" checked>
        7座 <input name="car15" type="checkbox" id="car15" value="15座">
        15座 
        <input name="car21" type="checkbox" id="car21" value="21座">
        21座 
        <input name="car25" type="checkbox" id="car25" value="25座">
        25座 
        <input name="car32" type="checkbox" id="car32" value="32座">
        32座 <font color="#CC0000">*</font></td>
    </tr>
    <tr> 
      <td height="167">
<div align="center">其它要求;</div></td>
      <td><textarea name="other_require" cols="50" rows="10" id="other_require"></textarea></td>
    </tr>
    <tr bgcolor="#FF9933" height="15"> 
      <td bgcolor="#cccccc" height="15">&nbsp;</td>
      <td bgcolor="#cccccc" height="15"> 
        <input type="submit" name="Submit" value="确 定">
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="reset" name="Submit2" value="重 写">
      </td>
    </tr>
  </table>
</form>
<div align="center">
  <script language="JScript">
<!--
function check(form1)
{
 var i
 i=0
if (form1.book_name.value=="")
 {
    alert("请输入您的姓名!")
	form1.book_name.focus()
    return false
  }
  else
  {
  i=i+1
  }
  
  if (form1.book_unit.value=="")
 {
    alert("请输入您的单位!")
	form1.book_unit.focus()
    return false
  }
  else
  {
  i=i+1
  }
  
  if (form1.contact_tel.value=="")
 {
    alert("请输入您的联系电话!")
	form1.contact_tel.focus()
    return false
  }
  else
  {
  i=i+1
  }
  
  if (form1.email.value=="")
 {
    alert("请输入您的E-mail!")
	form1.email.focus()
    return false
  }
  else
  {
  i=i+1
  }
  
  if (form1.book_date.value=="")
 {
    alert("请确定您的包车日期!")
	form1.book_date.focus()
    return false
  }
  else
  {
  i=i+1
  }
  
  if (form1.book_way.value=="")
 {
    alert("请确定您的包车路线!")
	form1.book_way.focus()
    return false
  }
  else
  {
  i=i+1
  }
  if(i==6)
  {
    return true
  }
}
//-->
</script>
  <a href="admin/index.asp">设置收信邮箱</a> </div>
</body>
</html>
