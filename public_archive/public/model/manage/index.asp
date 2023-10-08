<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Web Site Manage</title>
<%
 if session("admin_name")="" then response.Redirect("index.htm")
%>
</head>
<frameset cols="16%,80%" frameborder="yes" border="1" framespacing="0">
  <frame src="left.asp" name="left" noresize>
  <frame src="right.htm" name="main">
</frameset>
<noframes></noframes>
</html>
