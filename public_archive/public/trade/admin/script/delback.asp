<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Delete Back</title>
</head>
<body>
<%
id=request.QueryString("id")
sql="delete * from back where id="&id&""
conn.execute(sql)
%>
<script language="JavaScript">
alert("已成功册除该回复!")
location.href="viewback.asp"
</script>
</body>
</html>
