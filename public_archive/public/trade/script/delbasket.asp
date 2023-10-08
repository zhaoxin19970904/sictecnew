<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Delete Basket</title>
</head>

<body>
<%
username=session("username")
id=request.QueryString("id")
action=request.QueryString("action")
if action="all" then
  set rs=server.CreateObject("adodb.recordset")
  sql="delete * from basket where memberid='"&username&"'"
  rs.open sql,conn,1,3
%>
<script language="JavaScript">
alert("Successfully Delete!")
location.href="../viewbasket.asp"
</script>
<%
else
  sql="delete * from basket where id="&id&""
  conn.execute(sql)
%>
<script language="JavaScript">
alert("Successfully Delete!")
location.href="../viewbasket.asp"
</script>
<%
end if
%>
</body>
</html>
