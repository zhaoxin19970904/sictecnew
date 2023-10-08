<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>
<%
  id=request.QueryString("id")
  action=request.QueryString("action")
  if action="tj" then
     sql="update products set ifcommended=true where id="&id&""
     conn.execute(sql)
	 conn.close
     set conn=nothing
%>
<script language=javascript>
   alert("Successful!")
   location.href="viewpro.asp"
</script>

<%
   elseif action="cxtj" then
     sql="update products set ifcommended=false where id="&id&""
     conn.execute(sql)
	 conn.close
     set conn=nothing
%>
<script language=javascript>
   alert("Successful!")
   location.href="viewpro.asp"
</script>

<%
  end if  
%>
</body>
</html>
