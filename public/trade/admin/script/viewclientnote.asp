<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type="text/css">
.text{
  font-size:12px;
}
</style>
</head>
<body topmargin="50">
<%
table=request.QueryString("table")
id=request.QueryString("id")
sql="select notes from "&table&" where id="&id&""
set rs=conn.execute(sql)
%>
<table width="553" height="65" border="0" align="center" cellpadding="0" cellspacing="0" class="text">
  <tr>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("notes")%></td>
  </tr>
  <tr>
    <td height="49" valign="bottom"><div align="center">[ <a href="#" onclick="javascript:history.back()">返 回</a> 
        ]</div></td>
  </tr>
</table>
<%
rs.close
conn.close
set conn=nothing
%>
</body>
</html>
