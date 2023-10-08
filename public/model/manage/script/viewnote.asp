<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type=text/css>
.info{
FONT-SIZE:12px;
color:#999999;
}
a{
FONT-SIZE:12px;
}

</style>
</head>

<body>
<%
  id=request.QueryString("id")
  set rs=server.CreateObject("adodb.recordset")
  sql="select type,name,original,notes from modeldb where id="&id&""
  rs.open sql,conn,1,1
%>
<table width="606" border="0" align="center" cellpadding="0" cellspacing="0" class="info">
  <tr> 
    <td width="219" height="28"> <div align="center"><font color="#FF9900">Type</font></div></td>
    <td width="234"><div align="center"><font color="#FF9900">Name</font></div></td>
    <td width="153"><div align="center"><font color="#FF9900">Country</font></div></td>
  </tr>
  <tr> 
    <td height="32"><div align="center"><%=rs("type")%></div></td>
    <td><div align="center"><%=rs("name")%></div></td>
    <td><div align="center"><%=rs("original")%></div></td>
  </tr>
  <tr valign="bottom"> 
    <td height="45" colspan="3"> 
      <div align="center"><font color="#003366"><%=rs("notes")%></font></div></td>
  </tr>
</table>

<div align="center"><br>
  <br>
</div>
<div align="center"> [ <a href="#" onclick="javascript:history.back()">返 回</a> ] </div>
<%
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
%>
</body>
</html>
