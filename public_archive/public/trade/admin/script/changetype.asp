<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Change Type</title>
</head>
<body>
<%
del=request.Form("del")
type_id=request.Form("type_id")
newtype=rtrim(ltrim(request.Form("newclass")))
if del="" then
sql="select type from type where type_id="&type_id&""
set rs=conn.execute(sql)
old_type=rtrim(ltrim(rs(0)))
rs.close
sql="update type set type='"&newtype&"' where type='"&old_type&"'"
conn.execute(sql)
sql="update basket set type='"&newtype&"' where type='"&old_type&"'"
conn.execute(sql)
sql="update products set type='"&newtype&"' where type='"&old_type&"'"
conn.execute(sql)
%>
<script language="JavaScript">
alert("Successfully!")
location.href="addsmallclass.asp"
</script>
<%
else
sql="select id from products products where type_id="&type_id&""
set rs=conn.execute(sql)
  if not rs.eof then
  rs.close
%>
<script language="JavaScript">
alert("删除失败,请先在产品中删除该类的所有产品!")
location.href="addsmallclass.asp"
</script>
<%
else 
sql="delete * from type where type_id="&type_id&""
conn.execute(sql)
%>
<script language="JavaScript">
alert("Successfully!")
location.href="addsmallclass.asp"
</script>
<%end if%>
<%end if%>
</body>
</html>
