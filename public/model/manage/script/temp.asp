<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>
<%
'photo2=request.QueryString("photo2")
'if photo2<>"" then
 'else
photo2=session("photo2")
'end if
path=server.mappath("photo2/"&photo2)
set objfso = server.createobject("scripting.FileSystemObject")
     if objfso.FileExists(path) then
         objfso.DeleteFile(path)   
	     set objfso = nothing
		 end if
		 response.redirect("upload1.asp")
%>
</body>
</html>
