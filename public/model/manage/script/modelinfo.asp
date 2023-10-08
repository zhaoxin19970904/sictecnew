<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>模型的其它信息</title>
<script language="jscript">
<!--
  function check(form1)
  {
    var fig=true
	
    if(form1.m_Name.value=="")
	  {
	    window.alert("请输入模型名称！")
		fig=false
	  }
	  if(form1.original.value=="")
	  {
	    window.alert("请输入原创国家！")
		fig=false
	  }
	  if(form1.width_height.value=="")
	  {
	    window.alert("请确定图片的长宽比！")
		fig=false
	  }
	return fig
  }  
  
//-->
</script>


<style type=text/css>
.info{
FONT-SIZE:12px;
}
.input {
	border:1px solid #666666; padding:1px; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: "Verdana", "Arial", "宋体"; HEIGHT: 17px; BACKGROUND-COLOR: #ffffff
}
</style>
<%
max_sign=session("max_sign")
set rs=server.CreateObject("adodb.recordset")
sql="select photo1,photo2 from modeldb where sign="&max_sign&""
rs.open sql,conn,1,1
%>
</head>

<body topmargin="30">
<table width="637" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
  <tr>
    <td width="258" height="26"> 
      <div align="center">小图片名称:</div></td>
    <td width="18">&nbsp;</td>
    <td width="361"><div align="center">大图片名称:</div></td>
  </tr>
  <tr>
    <td><div align="center"><font color="0000ff">
	<a href="photo1/<%=rs("photo1")%>"><%=rs("photo1")%></a>
	</font></div></td>
    <td>&nbsp;</td>
    <td><div align="center"><font color="0000ff">
	<a href="photo2/<%=rs("photo2")%>"><%=rs("photo2")%></a>
	</font></div></td>
  </tr>
</table>

<div align="center"><br>
  [ <a href="endup.asp?deal=small&sign=<%=max_sign%>">结束上传</a> 
  ]<br>
</div>
<form name="form1" method="post" action="addmodel.asp" onsubmit="return check(this)">
  <table width="420" height="396" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
    <tr> 
      <td width="78"><div align="center">Type</div></td>
      <td width="73">(模型大类)</td>
      <td width="269"> <select name="m_Type" id="m_Type">
          <option selected>请选择大类</option>
          <option>Rocket</option>
          <option>Satellite</option>
		  <option>Spacecraft</option>
          <option>Spaceshuttle</option>
          <option>Airplane</option>
          <option>Space Station</option>
          <option>Lunar Module</option>
          <option>Others</option>
        </select> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">Name</div></td>
      <td>(模型名称)</td>
      <td><input name="m_Name" type="text" class="input" id="m_Name"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="18"> <div align="center">Manufacturer</div></td>
      <td>(制造厂商)</td>
      <td><input name="Manufacturer" type="text" class="input" id="Manufacturer"></td>
    </tr>
    <tr> 
      <td><div align="center">Country</div></td>
      <td>(原创国家)</td>
      <td><input name="original" type="text" class="input" id="original"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">Rato</div></td>
      <td>(比例)</td>
      <td><input name="Rato" type="text" class="input" id="Rato"></td>
    </tr>
    <tr> 
      <td><div align="center">Size</div></td>
      <td>(模型大小)</td>
      <td><input name="m_Size" type="text" class="input" id="m_Size"></td>
    </tr>
    <tr> 
      <td><div align="center">Org-size</div></td>
      <td>(实际大小)</td>
      <td><input name="Org-size" type="text" class="input" id="Org-size"></td>
    </tr>
    <tr> 
      <td><div align="center">Material</div></td>
      <td>(制作材料)</td>
      <td><input name="Material" type="text" class="input" id="Material"></td>
    </tr>
    <tr> 
      <td><div align="center">Price</div></td>
      <td>(价格)</td>
      <td><input name="Price" type="text" class="input" id="Price"></td>
    </tr>
    <tr> 
      <td><div align="center">Miniquantity</div></td>
      <td>(数量)</td>
      <td><input name="Miniquantity" type="text" class="input" id="Miniquantity"></td>
    </tr>
    <tr> 
      <td><div align="center">Leadtime</div></td>
      <td>&nbsp;</td>
      <td><input name="Leadtime" type="text" class="input" id="Leadtime"></td>
    </tr>
    <tr> 
      <td><div align="center">Pack</div></td>
      <td>&nbsp;</td>
      <td><input name="Packing" type="text" class="input" id="Packing"></td>
    </tr>
    <tr> 
      <td><div align="center">Sample</div></td>
      <td>(有无样品)</td>
      <td><select name="Sample" id="Sample">
          <option selected>有</option>
          <option>无</option>
        </select></td>
    </tr>
    <tr> 
      <td height="25" colspan="3"><div align="center">
          <input type="radio" name="width_height" value="widthW">
          模型图片的<font color="#000000">长</font><font color="#FF0000">大于</font>宽&nbsp; 
          <a href="example.asp?id=2" target="_blank">如图示样</a> &nbsp;&nbsp; 
          <input type="radio" name="width_height" value="heightH">
          模型图片的<font color="#000000">长</font><font color="#FF0000">小于<font color="#000000">宽 
          <a href="example.asp?id=2" target="_blank">如图示样</a> </font></font><font color="#FF0000">*</font></div></td>
    </tr>
    <tr> 
      <td><div align="center">Note</div></td>
      <td>(备注)</td>
      <td><textarea name="Note" id="Note"></textarea></td>
    </tr>
    <tr> 
      <td height="48" valign="bottom">&nbsp;</td>
      <td>&nbsp;</td>
      <td valign="bottom"> <input type="submit" name="Submit" value="确定上传" class="input"> 
        &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; <input type="reset" name="Submit2" value="重写信息" class="input"> 
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
