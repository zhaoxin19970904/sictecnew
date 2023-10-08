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
id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
sql="select * from modeldb where id="&id&""
rs.open sql,conn,1,1
%>
</head>

<body topmargin="30">
<div align="center">
  <br>
    <font color="#666666">模型信息修改</font><br>
 <font color="#666666"><img src="../../images/line_01.jpg" width="338" height="20"></font>
</div>
<form name="form1" method="post" action="saveedit.asp?id=<%=rs("id")%>" onsubmit="return check(this)">
  <table width="419" height="390" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
    <tr> 
      <td width="78"><div align="center">Type</div></td>
      <td width="73">(模型大类)</td>
      <td width="268"> <select name="m_Type" id="m_Type">
          <option <%if rs("type")="rocket" then response.write "selected" else response.write "" end if%>>rocket</option>
          <option <%if rs("type")="satellite" then response.write "selected" else response.write "" end if%>>satellite</option>
          <option <%if rs("type")="Spacecraft" then response.write "selected" else response.write "" end if%>>Spacecraft</option>
          <option <%if rs("type")="spaceshuttle" then response.write "selected" else response.write "" end if%>>spaceshuttle</option>
          <option <%if rs("type")="airplane" then response.write "selected" else response.write "" end if%>>airplane</option>
          <option <%if rs("type")="space station" then response.write "selected" else response.write "" end if%>>space 
          station</option>
          <option <%if rs("type")="lunar module" then response.write "selected" else response.write "" end if%>>lunar 
          module</option>
          <option <%if rs("type")="others" then response.write "selected" else response.write "" end if%>>others</option>
        </select> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">Name</div></td>
      <td>(模型名称)</td>
      <td><input name="m_Name" type="text" class="input" id="m_Name" value="<%=rs("Name")%>"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="18"> <div align="center">Manufacturer</div></td>
      <td>(制造厂商)</td>
      <td> <input name="Manufacturer" type="text" class="input" id="Manufacturer" value="<%=rs("Manufacturer")%>"> 
      </td>
    </tr>
    <tr> 
      <td><div align="center">Country</div></td>
      <td>(原创国家)</td>
      <td><input name="original" type="text" class="input" id="original" value="<%=rs("original")%>"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">Rato</div></td>
      <td>(比例)</td>
      <td> <input name="Rato" type="text" class="input" id="Rato" value="<%=rs("Ratio")%>"> 
      </td>
    </tr>
    <tr> 
      <td><div align="center">Size</div></td>
      <td>(模型大小)</td>
      <td> <input name="m_Size" type="text" class="input" id="m_Size" value="<%=rs("Size")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">Org-size</div></td>
      <td>(实际大小)</td>
      <td> <input name="Org-size" type="text" class="input" id="Org-size" value="<%=rs("Org-size")%>"> 
      </td>
    </tr>
    <tr> 
      <td><div align="center">Material</div></td>
      <td>(制作材料)</td>
      <td><input name="Material" type="text" class="input" id="Material" value="<%=rs("Material")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">Price</div></td>
      <td>(价格)</td>
      <td> <input name="Price" type="text" class="input" id="Price" value="<%=rs("Price")%>"> 
      </td>
    </tr>
    <tr> 
      <td><div align="center">Miniquantity</div></td>
      <td>(数量)</td>
      <td><input name="Miniquantity" type="text" class="input" id="Miniquantity" value="<%=rs("Miniquantity")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">Leadtime</div></td>
      <td>&nbsp;</td>
      <td><input name="Leadtime" type="text" class="input" id="Leadtime" value="<%=rs("Leadtime")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">Pack</div></td>
      <td>&nbsp;</td>
      <td><input name="Packing" type="text" class="input" id="Packing" value="<%=rs("Packing")%>"> 
      </td>
    </tr>
    <tr> 
      <td><div align="center">Sample</div></td>
      <td>(有无样品)</td>
      <td><select name="Sample" id="Sample">
          <option <%if rs("Sample")=true then response.write "selected" else response.write "" end if%>>有</option>
          <option <%if rs("Sample")=false then response.write "selected" else response.write "" end if%>>无</option>
        </select></td>
    </tr>
    <tr> 
      <td height="19" colspan="3">
<div align="center"> 
          <input name="width_height" type="radio" value="widthW" <%if rs("w_h")="widthW" then response.Write("checked") else response.Write "" end if%>>
          模型图片的<font color="#000000">长</font><font color="#FF0000">大于</font>宽 
          &nbsp;<a href="example.asp?id=2" target="_blank">如图示样</a>&nbsp; 
          <input type="radio" name="width_height" value="heightH" <%if rs("w_h")="heightH" then response.Write("checked") else response.Write "" end if%>>
          模型图片的<font color="#000000">长</font><font color="#FF0000">小于<font color="#000000">宽 
          <a href="example.asp?id=1" target="_blank">如图示样</a> </font>* </font></div></td>
    </tr>
    <tr> 
      <td><div align="center">Note</div></td>
      <td>(备注)</td>
      <td><textarea name="Note" id="Note"><%=rs("Notes")%></textarea></td>
    </tr>
    <tr> 
      <td height="48" valign="bottom">&nbsp;</td>
      <td colspan="2" valign="bottom"> <input type="submit" name="Submit" value="确定修改" class="input"> 
        &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; <input type="button" name="Submit2" value="取消修改" class="input" onclick="javascript:history.back()"> 
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
