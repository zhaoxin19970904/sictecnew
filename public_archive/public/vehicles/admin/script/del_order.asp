<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Delete Order</title>
</head>
<body>
<%
  id=request.querystring("id")
  sql="delete * from order_ where id="&id&""
  conn.execute(sql)
%>
<script language="javascript">
  alert("Successfully Delete!")
  location.href="vieworder.asp"
</script>
</body>
</html>
