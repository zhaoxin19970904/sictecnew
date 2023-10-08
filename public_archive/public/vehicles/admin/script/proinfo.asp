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
    <td height="26"> <div align="center">The product name:</div>
      <div align="center"></div></td>
  </tr>
  <tr> 
    <td> <div align="center"><font color="0000ff"> 
	<a href="../../prophotos/<%=filename%>" target="_blank"><%=filename%></a> </font></div>
      <div align="center"><font color="0000ff"> </font></div></td>
  </tr>
</table>
<div align="center"><br>
  [ <a href="endup.asp?filename=<%=filename%>">End up product</a> 
  ]<br>
</div>
<form name="form1" method="post" action="addpro.asp" onsubmit="return check(this)">
  <table width="538" height="324" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
    <tr> 
      <td width="164" height="59"> <div align="center">type:</div></td>
      <td width="10">&nbsp;</td>
      <td width="364"> <%
		sql="select distinct type_code,type from type"
		set rs=conn.execute(sql)
	  %> <select name="type" id="type">
          <option value="" selected>==select==</option>
          <%
		do while not rs.eof
		%>
          <option value="<%=rs(0)%>"><%=rs(1)%></option>
          <%
		 if not rs.eof then rs.movenext
		 loop
		 %>
        </select> <%
		 rs.close
		 set rs=nothing
		 conn.close
		 set conn=nothing
		%> <font color="#FF0000">* 
        <input type="hidden" name="hiddenField">
        </font></td>
    </tr>
    <tr> 
      <td height="23">
<div align="center">model:</div></td>
      <td>&nbsp;</td>
      <td><input name="model" type="text" class="input" id="model"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="31"><div align="center">price:</div></td>
      <td>&nbsp;</td>
      <td><input name="price" type="text" id="price" size="10">
        &yen;</td>
    </tr>
    <tr> 
      <td height="31"><div align="center">if commended:</div></td>
      <td>&nbsp;</td>
      <td><input type="radio" name="ifcommended" value="yes">
        yes 
        <input name="ifcommended" type="radio" value="no" checked>
        no</td>
    </tr>
    <tr> 
      <td><div align="center">Introduction:</div></td>
      <td>&nbsp;</td>
      <td><textarea name="Note" cols="30" rows="8" id="Note"></textarea></td>
    </tr>
    <tr> 
      <td height="48" valign="bottom">&nbsp;</td>
      <td>&nbsp;</td>
      <td valign="bottom"> <input type="submit" name="Submit" value="Submit" class="input" onclick="javascript:return checkempty(this.form); "> 
        &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; <input type="reset" name="Submit2" value="Reset" class="input"> 
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
