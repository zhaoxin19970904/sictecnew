<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ģ�͵�������Ϣ</title>
<script language="jscript">
<!--
  function check(form1)
  {
    var fig=true
	
    if(form1.m_Name.value=="")
	  {
	    window.alert("������ģ�����ƣ�")
		fig=false
	  }
	  if(form1.original.value=="")
	  {
	    window.alert("������ԭ�����ң�")
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
	border:1px solid #666666; padding:1px; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: "Verdana", "Arial", "����"; HEIGHT: 17px; BACKGROUND-COLOR: #ffffff
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
    <font color="#666666">ģ����Ϣ�޸�</font><br>
 <font color="#666666"><img src="../../images/line_01.jpg" width="338" height="20"></font>
</div>
<form name="form1" method="post" action="saveedit.asp?id=<%=rs("id")%>" onsubmit="return check(this)">
  <table width="419" height="390" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
    <tr> 
      <td width="78"><div align="center">Type</div></td>
      <td width="73">(ģ�ʹ���)</td>
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
      <td>(ģ������)</td>
      <td><input name="m_Name" type="text" class="input" id="m_Name" value="<%=rs("Name")%>"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="18"> <div align="center">Manufacturer</div></td>
      <td>(���쳧��)</td>
      <td> <input name="Manufacturer" type="text" class="input" id="Manufacturer" value="<%=rs("Manufacturer")%>"> 
      </td>
    </tr>
    <tr> 
      <td><div align="center">Country</div></td>
      <td>(ԭ������)</td>
      <td><input name="original" type="text" class="input" id="original" value="<%=rs("original")%>"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">Rato</div></td>
      <td>(����)</td>
      <td> <input name="Rato" type="text" class="input" id="Rato" value="<%=rs("Ratio")%>"> 
      </td>
    </tr>
    <tr> 
      <td><div align="center">Size</div></td>
      <td>(ģ�ʹ�С)</td>
      <td> <input name="m_Size" type="text" class="input" id="m_Size" value="<%=rs("Size")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">Org-size</div></td>
      <td>(ʵ�ʴ�С)</td>
      <td> <input name="Org-size" type="text" class="input" id="Org-size" value="<%=rs("Org-size")%>"> 
      </td>
    </tr>
    <tr> 
      <td><div align="center">Material</div></td>
      <td>(��������)</td>
      <td><input name="Material" type="text" class="input" id="Material" value="<%=rs("Material")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">Price</div></td>
      <td>(�۸�)</td>
      <td> <input name="Price" type="text" class="input" id="Price" value="<%=rs("Price")%>"> 
      </td>
    </tr>
    <tr> 
      <td><div align="center">Miniquantity</div></td>
      <td>(����)</td>
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
      <td>(������Ʒ)</td>
      <td><select name="Sample" id="Sample">
          <option <%if rs("Sample")=true then response.write "selected" else response.write "" end if%>>��</option>
          <option <%if rs("Sample")=false then response.write "selected" else response.write "" end if%>>��</option>
        </select></td>
    </tr>
    <tr> 
      <td height="19" colspan="3">
<div align="center"> 
          <input name="width_height" type="radio" value="widthW" <%if rs("w_h")="widthW" then response.Write("checked") else response.Write "" end if%>>
          ģ��ͼƬ��<font color="#000000">��</font><font color="#FF0000">����</font>�� 
          &nbsp;<a href="example.asp?id=2" target="_blank">��ͼʾ��</a>&nbsp; 
          <input type="radio" name="width_height" value="heightH" <%if rs("w_h")="heightH" then response.Write("checked") else response.Write "" end if%>>
          ģ��ͼƬ��<font color="#000000">��</font><font color="#FF0000">С��<font color="#000000">�� 
          <a href="example.asp?id=1" target="_blank">��ͼʾ��</a> </font>* </font></div></td>
    </tr>
    <tr> 
      <td><div align="center">Note</div></td>
      <td>(��ע)</td>
      <td><textarea name="Note" id="Note"><%=rs("Notes")%></textarea></td>
    </tr>
    <tr> 
      <td height="48" valign="bottom">&nbsp;</td>
      <td colspan="2" valign="bottom"> <input type="submit" name="Submit" value="ȷ���޸�" class="input"> 
        &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; <input type="button" name="Submit2" value="ȡ���޸�" class="input" onclick="javascript:history.back()"> 
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
