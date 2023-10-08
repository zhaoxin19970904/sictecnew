<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>模型的其它信息</title>

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
%>
</head>

<body topmargin="30">
<div align="center"><font color="#666666">产品信息修改</font><br>
  <font color="#666666"><img src="../../images/line_01.jpg" width="338" height="20"></font> 
</div>
<form name="form1" method="post" action="saveditproinfo.asp?id=<%=id%>">
  <table width="538" height="484" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
    <tr> 
      <td width="164"><div align="center">type</div></td>
      <td width="10">&nbsp;</td>
      <td width="364"> 
	  <%
		'sql="select distinct main_id,type from type"
		sql="select distinct type_id,type from type"
		sql1="select * from products where id="&id&""
		set rs=conn.execute(sql)
		set rs1=conn.execute(sql1)
	  %>
	  <select name="type" id="type">
		<%do while not rs.eof%>
          <option value="<%=rs(0)%>:<%=rs(1)%>" <%if rs1("type")=rs(1) then response.Write "selected" else response.Write "" end if%>><%=rs(1)%></option>
         <%
		 if not rs.eof then rs.movenext
		 loop
		 %> 
        </select>
		<%
		 rs.close
		 set rs=nothing
		%>
		 <font color="#FF0000">*</font></td>
    </tr>
    
    <tr> 
      <td><div align="center">name</div></td>
      <td>&nbsp;</td>
      <td><input name="name" type="text" class="input" id="name"  value="<%=rs1("name")%>"><font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td><div align="center">model</div></td>
      <td>&nbsp;</td>
      <td><input name="model" type="text" class="input" id="model" value="<%=rs1("model")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">brandname</div></td>
      <td>&nbsp;</td>
      <td><input name="brandname" type="text" class="input" id="brandname" value="<%=rs1("brandname")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">region</div></td>
      <td>&nbsp;</td>
      <td><input name="region" type="text" class="input" id="region" value="<%=rs1("region")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">approvals</div></td>
      <td>&nbsp;</td>
      <td><input name="approvals" type="text" class="input" id="approvals" value="<%=rs1("approvals")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">Price</div></td>
      <td>&nbsp;</td>
      <td><input name="Price" type="text" class="input" id="Price" value="<%=rs1("Price")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">main_export_markets</div></td>
      <td>&nbsp;</td>
      <td><input name="main_export_markets" type="text" class="input" id="main_export_markets"  value="<%=rs1("main_export_markets")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">Packing</div></td>
      <td>&nbsp;</td>
      <td><input name="Packing" type="text" class="input" id="Packing" value="<%=rs1("Packing")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">key_specification</div></td>
      <td>&nbsp;</td>
      <td><input name="key_specification" type="text" class="input" id="key_specification" value="<%=rs1("key_specification")%>"></td>
    </tr>
    <tr> 
      <td height="22"><div align="center">payment_terms</div></td>
      <td>&nbsp;</td>
      <td><input name="payment_terms" type="text" id="payment_terms" class="input" value="<%=rs1("payment_terms")%>"></td>
    </tr>
    <tr> 
      <td height="19"><div align="center">lead_time</div></td>
      <td>&nbsp;</td>
      <td><input name="lead_time" type="text" id="lead_time" class="input" value="<%=rs1("lead_time")%>"></td>
    </tr>
    <tr> 
      <td height="21"><div align="center">cost</div></td>
      <td>&nbsp;</td>
      <td><input name="cost" type="text" id="cost" class="input" value="<%=rs1("cost")%>"></td>
    </tr>
    <tr> 
      <td height="19"><div align="center">manufacture</div></td>
      <td>&nbsp;</td>
      <td><input name="manufacture" type="text" id="manufacture" class="input" value="<%=rs1("manufacture")%>"></td>
    </tr>
    <tr> 
      <td height="22"><div align="center">related_products</div></td>
      <td>&nbsp;</td>
      <td><input name="related_products" type="text" id="related_products" class="input" value="<%=rs1("related_products")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">key_words</div></td>
      <td>&nbsp;</td>
      <td><input name="key_words" type="text" id="key_words" class="input" value="<%=rs1("key_words")%>"></td>
    </tr>
    <tr> 
      <td><div align="center">Sample</div></td>
      <td>&nbsp;</td>
      <td>
	  <select name="sample" id="sample">
          <option value="有" <%if rs1("sample")=true then response.write "selected" else response.write "" end if%>>有</option>
          <option value="无" <%if rs1("sample")=false then response.write "selected" else response.write "" end if%>>无</option>
      </select>
		</td>
    </tr>
    <tr> 
      <td><div align="center">Note</div></td>
      <td>&nbsp;</td>
      <td><textarea name="Note" id="Note"><%=rs1("Notes")%></textarea></td>
    </tr>
    <tr> 
      <td height="48" valign="bottom">&nbsp;</td>
      <td>&nbsp;</td>
      <td valign="bottom"> 
	      <input type="submit" name="Submit" value="确定上传" class="input" onclick="javascript:return checkempty(this.form); "> 
        &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;
		 <input type="button" name="Submit2" value="放弃修改" class="input" onclick="javascript:history.back();"> 
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <%
  rs1.close
  set rs1=nothing
  conn.close
  set conn=nothing
  %>
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
