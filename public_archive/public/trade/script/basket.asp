<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Add To Basdet</title>
</head>
<body>
<%
username=session("username")
if username="" then
%>
<script language="JavaScript">
alert("Please login now!")
location.href="../signin.asp"
</script>

<%
else
productno=request.QueryString("pro_id")
sql="select name,photo from products where id="&productno&""
set rs=conn.execute(sql)
pro_name=rs("name")
photo=rs("photo")
rs.close
type_id=request.QueryString("type_id")
sql="select type from type where type_id="&type_id&""
set rs=conn.execute(sql)
type_type=rs(0)
rs.close
set rs=server.CreateObject("adodb.recordset")
sql="select * from basket"
rs.open sql,conn,1,3
rs.addnew
rs("memberid")=username
rs("productno")=productno
rs("photo")=photo
rs("pro_name")=pro_name
rs("type")=type_type
rs("date")=now()
rs("book_ip")=request.ServerVariables("REMOTE_ADDR")
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script language="JavaScript">
alert("Successfully!")
location.href="../viewbasket.asp"
</script>
<%
end if
%>
</body>
</html>
