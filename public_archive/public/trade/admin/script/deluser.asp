<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Delete User</title>
</head>

<body>
<%
id=request.QueryString("id")
sql="delete * from member where id="&id&""
conn.execute(sql)
%>
<script language="JavaScript">
alert("Successfully!")
location.href="usermanage.asp"
</script>
</body>
</html>
