<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<body>
<%
table=request.QueryString("table")
id=request.QueryString("id")
sql="delete * from "&table&" where id="&id&""
conn.execute(sql)
conn.close
set conn=nothing
%>
<script language=javascript>
   alert("Successful!")
</script>
<%
  response.Redirect "view"&table&".asp"
%>
</body>
</html>
