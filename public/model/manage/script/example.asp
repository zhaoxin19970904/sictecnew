<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ͼƬʾ��</title>
</head>
<body>
<%
  id=request.QueryString("id")
  if id=2 then
%>
<img src="../../images/mo_4.jpg">
<%else%>
<img src="../../images/mo_3.jpg">
<%end if%>
</body>
</html>
