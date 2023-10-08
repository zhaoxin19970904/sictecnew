<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Delete Write Back</title>
</head>
<body>
<%
action=request.QueryString("action")
if action="all" then
  username=session("username")
  set rs=server.CreateObject("adodb.recordset")
  sql="delete * from back where ask_member='"&username&"'"
  rs.open sql,conn,1,3
%>
<script>
 alert("Successfully delete all!")
 location.href="../view_back.asp"
</script>
<%
else
id=request.QueryString("id")
sql="delete * from back where id="&id&""
conn.execute(sql)
%>
<script>
 alert("Successfully delete!")
 location.href="../view_back.asp"
</script>
<%end if%>
</body>
</html>
