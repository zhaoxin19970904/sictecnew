<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Information</title>

<style type=text/css>
.info{
FONT-SIZE:12px;
}
.input {
	border:1px solid #666666; padding:1px; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: "Verdana", "Arial", "宋体"; HEIGHT: 17px; BACKGROUND-COLOR: #ffffff
}
</style>
<%
filename=session("filename")
set rs=server.CreateObject("adodb.recordset")
%>
</head>

<body topmargin="30">
<table width="637" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
  <tr> 
    <td height="26"> <div align="center">产品图片名称:</div>
      <div align="center"></div></td>
  </tr>
  <tr> 
    <td> <div align="center"><font color="0000ff"> <a href="prophotos/<%=filename%>" target="_blank"><%=filename%></a> </font></div>
      <div align="center"><font color="0000ff"> </font></div></td>
  </tr>
</table>
<div align="center"><br>
  [ <a href="endup.asp?filename=<%=filename%>">结束上传</a> 
  ]<br>
</div>
<form name="form1" method="post" action="addpro.asp" onsubmit="return check(this)">
  <table width="538" height="484" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
    <tr> 
      <td width="164"><div align="center">type</div></td>
      <td width="10">&nbsp;</td>
      <td width="364"> 
	  <%
		sql="select distinct type_id,type from type"
		set rs=conn.execute(sql)
	  %>
	  <select name="type" id="type">
          <option value="" selected>==select==</option>
		<%
		do while not rs.eof
		%>
          <option value="<%=rs(0)%>:<%=rs(1)%>"><%=rs(1)%></option>
         <%
		 if not rs.eof then rs.movenext
		 loop
		 %>
        </select>
		<%
		 rs.close
		 set rs=nothing
		 conn.close
		 set conn=nothing
		%>
        <font color="#FF0000">*
        <input type="hidden" name="hiddenField">
        </font></td>
    </tr>
    
    <tr> 
      <td><div align="center">name</div></td>
      <td>&nbsp;</td>
      <td><input name="name" type="text" class="input" id="name"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">model</div></td>
      <td>&nbsp;</td>
      <td><input name="model" type="text" class="input" id="model"></td>
    </tr>
    <tr> 
      <td><div align="center">brandname</div></td>
      <td>&nbsp;</td>
      <td><input name="brandname" type="text" class="input" id="brandname"></td>
    </tr>
    <tr> 
      <td><div align="center">region</div></td>
      <td>&nbsp;</td>
      <td><input name="region" type="text" class="input" id="region"></td>
    </tr>
    <tr> 
      <td><div align="center">approvals</div></td>
      <td>&nbsp;</td>
      <td><input name="approvals" type="text" class="input" id="approvals"></td>
    </tr>
    <tr> 
      <td><div align="center">Price</div></td>
      <td>&nbsp;</td>
      <td><input name="Price" type="text" class="input" id="Price"></td>
    </tr>
    <tr> 
      <td><div align="center">main_export_markets</div></td>
      <td>&nbsp;</td>
      <td><input name="main_export_markets" type="text" class="input" id="main_export_markets"></td>
    </tr>
    <tr> 
      <td><div align="center">Packing</div></td>
      <td>&nbsp;</td>
      <td><input name="Packing" type="text" class="input" id="Packing"></td>
    </tr>
    <tr> 
      <td><div align="center">key_specification</div></td>
      <td>&nbsp;</td>
      <td><input name="key_specification" type="text" class="input" id="key_specification"></td>
    </tr>
    <tr> 
      <td height="22"><div align="center">payment_terms</div></td>
      <td>&nbsp;</td>
      <td><input name="payment_terms" type="text" id="payment_terms" class="input"></td>
    </tr>
    <tr> 
      <td height="19"><div align="center">lead_time</div></td>
      <td>&nbsp;</td>
      <td><input name="lead_time" type="text" id="lead_time" class="input"></td>
    </tr>
    <tr> 
      <td height="21"><div align="center">cost</div></td>
      <td>&nbsp;</td>
      <td><input name="cost" type="text" id="cost" class="input"></td>
    </tr>
    <tr> 
      <td height="19"><div align="center">manufacture</div></td>
      <td>&nbsp;</td>
      <td><input name="manufacture" type="text" id="manufacture" class="input"></td>
    </tr>
    <tr> 
      <td height="22"><div align="center">related_products</div></td>
      <td>&nbsp;</td>
      <td><input name="related_products" type="text" id="related_products" class="input"></td>
    </tr>
    <tr> 
      <td><div align="center">key_words</div></td>
      <td>&nbsp;</td>
      <td><input name="key_words" type="text" id="key_words" class="input"></td>
    </tr>
    <tr> 
      <td><div align="center">Sample</div></td>
      <td>&nbsp;</td>
      <td><select name="Sample" id="Sample">
          <option value="有">有</option>
          <option value="无">无</option>
        </select></td>
    </tr>
    <tr> 
      <td><div align="center">Note</div></td>
      <td>&nbsp;</td>
      <td><textarea name="Note" id="Note"></textarea></td>
    </tr>
    <tr> 
      <td height="48" valign="bottom">&nbsp;</td>
      <td>&nbsp;</td>
      <td valign="bottom"> 
	      <input type="submit" name="Submit" value="确定上传" class="input" onclick="javascript:return checkempty(this.form); "> 
        &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;
		 <input type="reset" name="Submit2" value="重写信息" class="input"> 
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<script language=javascript>
  function checkempty(form1)
  {
    
	if(form1.type[form1.type.selectedIndex].value == "")
		{
			alert("Please select type!");
			form1.type.focus();
			return false;
		}
	
    if(form1.name.value=="")
	  {
	    alert("Please input product name！");
		form1.name.focus();
		return false;
	  }
  }  
</script>
</body>
</html>
