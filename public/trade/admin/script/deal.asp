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
     sql="update products set recommended=true where id="&id&""
     conn.execute(sql)
	 conn.close
     set conn=nothing
%>
<script language=javascript>
   alert("Successful!")
   location.href="viewpro.asp"
</script>
<%
   elseif action="ts" then
     sql="update products set especially=true where id="&id&""
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
     sql="update products set recommended=false where id="&id&""
     conn.execute(sql)
	 conn.close
     set conn=nothing
%>
<script language=javascript>
   alert("Successful!")
   location.href="viewpro.asp"
</script>
<%
   elseif action="cxts" then
     sql="update products set especially=false where id="&id&""
     conn.execute(sql)
	 conn.close
     set conn=nothing
%>
<script language=javascript>
   alert("Successful!")
   location.href="viewpro.asp"
</script>


<%
   elseif action="cxnew" then
     sql="update products set ifnew=false where id="&id&""
     conn.execute(sql)
	 conn.close
     set conn=nothing
%>
<script language=javascript>
   alert("Successful!")
   location.href="viewpro.asp"
</script>
<%
   elseif action="isnew" then
     sql="update products set ifnew=true where id="&id&""
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
