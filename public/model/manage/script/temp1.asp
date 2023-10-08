<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>
<%
'photo1=request.QueryString("photo1")
'if photo1<>"" then
' else
 photo1=session("photo1")
'end if
path=server.mappath("photo1/"&photo1)
set objfso = server.createobject("scripting.FileSystemObject")
     if objfso.FileExists(path) then
         objfso.DeleteFile(path)   
	     set objfso = nothing
		 end if
		 response.Redirect "viewmodel.asp"
%>
</body>
</html>
