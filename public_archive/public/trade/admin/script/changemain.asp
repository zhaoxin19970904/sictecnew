<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Change Main</title>
</head>

<body>
<%
del=request.Form("del")
main_id=request.Form("main_id")
newclass=request.Form("newclass2")
if del="" then
sql="update main set main_type='"&newclass&"' where main_id="&main_id&""
conn.execute(sql)
%>
<script language="JavaScript">
alert("Successfully!")
location.href="addallclass.asp"
</script>
<%
else
sql="delete * from main where main_id="&main_id&""
conn.execute(sql)
%>
<script language="JavaScript">
alert("Successfully!")
location.href="addallclass.asp"
</script>
<%
end if
%>
</body>
</html>
