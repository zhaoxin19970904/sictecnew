<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>
<%
options=request.form("options")
id=request.Form("subject")
action=request.QueryString("action")

if options="del" then
  sql="select title_type from title where title_id="&id&""
    set rs=conn.execute(sql)
	title_type=rs(0)
	rs.close
	sql="select id from products where title_type='"&title_type&"'"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof then
         sql="delete * from title where title_id="&id&""
         conn.execute(sql)
%>
<script language=javascript>
   alert("Successful!")
   location.href="addbclass.asp"
</script>
    <%else%>
<script language=javascript>
   alert("在产品信息中存在该类的产品,\n\n请先删除该类的所有产品信息后再删除该类!")
   location.href="addbclass.asp"
</script>
	<%end if%>
    
<%
  elseif options="rename" then
    sql="select title_type from title where title_id="&id&""
    set rs=conn.execute(sql)
	title_type=rs(0)
	rs.close
  
   retitle=request.Form("retitle")
    sql="update products set title_type='"&retitle&"' where title_type='"&title_type&"'"
	conn.execute(sql)
   sql="update title set title_type='"&retitle&"' where title_id="&id&""
   conn.execute(sql)
%>
<script language=javascript>
   alert("Successful!")
   location.href="addbclass.asp"
</script>
<%
  elseif options="new" then
    newtitle=request.Form("newtitle")
	set rs=server.CreateObject("adodb.recordset")
    sql="select * from title where title_id="&id&""
   rs.open sql,conn,1,3
   rs.addnew
   rs("title_type")=newtitle
   rs.update
   rs.close
   set rs=nothing
%>
<script language=javascript>
   alert("Successful!")
   location.href="addbclass.asp"
</script>
<%
end if
if action="bigclass" then
  set rs=server.CreateObject("adodb.recordset")
    sql="select * from main"
   rs.open sql,conn,1,3
   rs.addnew
   rs("title_id")=id
   rs("main_type")=trim(request.Form("newclass"))
   rs.update
   rs.close
   set rs=nothing
%>
<script language=javascript>
   alert("Successful!")
   location.href="addbclass.asp"
</script>

<%
elseif action="smallclass" then
  set rs=server.CreateObject("adodb.recordset")
    sql="select * from type"
   rs.open sql,conn,1,3
   rs.addnew
   rs("main_id")=id
   rs("type")=trim(request.Form("newclass"))
   rs.update
   rs.close
   set rs=nothing
%>
<script language=javascript>
   alert("Successful!")
   location.href="addsmallclass.asp"
</script>

<%end if%>
<%
 conn.close
 set conn=nothing
%>
</body>
</html>
