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
dim type_code(),pv_type(),i
i=1
id=request.QueryString("id")
'set rs=server.CreateObject("adodb.recordset")
sql="select * from products where id="&id&""
set rs=conn.execute(sql)
modelno=rs("modelno")
price=rs("price")
p_type=trim(rs("type"))
intro=rs("introduction")
ifcommended=rs("ifcommended")
introduction=rs("introduction")
rs.close
sql="select type_code,type from type"
set rs=conn.execute(sql)
do while not rs.eof 
   redim preserve type_code(i),pv_type(i)
    type_code(i)=rs("type_code")
	pv_type(i)=rs("type")
  if not rs.eof then rs.movenext
  i=i+1
loop
rs.close
pvs=ubound(type_code)
%>
</head>
<body topmargin="30">
<div align="center"><font color="#666666">Edit Product Information</font><br>
  <font color="#666666"><img src="../../images/line_01.jpg" width="338" height="20"></font> 
  <br>
</div>
<form name="form1" method="post" action="saveditproinfo.asp" onsubmit="return check(this)">
  <table width="400" height="331" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
    <tr> 
      <td width="164" height="59"> <div align="center">type:</div></td>
      <td width="10">&nbsp;</td>
      <td width="364"> <select name="type" id="type">
          <%
		   for i=1 to pvs
		  %>
          <option value="<%=type_code(i)%>" <%if type_code(i)=p_type then response.Write "selected" else response.Write "" end if%>><%=pv_type(i)%></option>
          <%
           next
		 %>
        </select> <%
		 conn.close
		 set conn=nothing
		%> <font color="#FF0000">* 
        <input type="hidden" name="product_id" value="<%=id%>">
        </font></td>
    </tr>
    <tr> 
      <td height="30">
<div align="center">model:</div></td>
      <td>&nbsp;</td>
      <td><input name="model" type="text" class="input" id="model" value="<%=modelno%>"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="31"><div align="center">price:</div></td>
      <td>&nbsp;</td>
      <td><input name="price" type="text" id="price" size="10" value="<%=price%>"> &yen;</td>
    </tr>
    <tr> 
      <td height="31"><div align="center">if commended:</div></td>
      <td>&nbsp;</td>
      <td><input type="radio" name="ifcommended" value="yes" <%if ifcommended=true then response.Write "checked" else response.Write "" end if%>>
        yes 
        <input name="ifcommended" type="radio" value="no" <%if ifcommended=false then response.Write "checked" else response.Write "" end if%>>
        no</td>
    </tr>
    <tr> 
      <td><div align="center">Introduction:</div></td>
      <td>&nbsp;</td>
      <td> <textarea name="Note" cols="30" rows="8" id="Note"><%=introduction%></textarea> 
      </td>
    </tr>
    <tr> 
      <td height="48" valign="bottom">&nbsp;</td>
      <td>&nbsp;</td>
      <td valign="bottom"> <input type="submit" name="Submit" value="Submit" class="input" onclick="javascript:return checkempty(this.form); "> 
        &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; <input type="button" name="Submit2" value="Give up " class="input" onClick="javascript:history.back();"> 
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
	  
	  if(form1.model.value=="")
	  {
	    alert("Please input product model！");
		form1.model.focus();
		return false;
	  }
  }  
</script>
</body>
</html>
