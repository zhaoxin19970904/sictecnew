<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>结束上传模型</title>
</head>
<body>
<%
On error resume next
filename=trim(request.QueryString("filename"))
id=trim(request.QueryString("id"))
if filename<>"" then
set rs=server.createobject("adodb.recordset")
path=server.mappath("../../proimages/"&filename)
strsql = "delete * from products where photo='"&filename&"'"
conn.execute strsql,intno
if intno <> 0 then
   strstring = "<li>The data is deleted successfully!"
   set objfso = server.createobject("scripting.FileSystemObject")
     if objfso.FileExists(path) then
         objfso.DeleteFile(path)   
	     set objfso = nothing
	    strstring = strstring &"<li>The file is deleted successfully from hard disk!"
                else
                       response.write "<script language='Javascript'>alert('The file have not found from hard disk ,it maybe deleted!')</script>"	      
                end if
else
strstring = "<li>Parameter is wrong!"
end if
set rs=nothing
conn.close
set conn = nothing
response.write strstring
response.write "<li>After two seconds system is auto back …"
%>
<meta http-equiv="refresh" content='2;url=viewpro.asp'>
<%

else
set rs=conn.execute("select photo from products where id="&id&"")
path=server.mappath("../../proimages/"&rs(0))
strsql = "delete * from products where id="&id&""
conn.execute strsql,intno
if intno <> 0 then
   strstring = "<li>The data is deleted successfully!"
   set objfso = server.createobject("scripting.FileSystemObject")
     if objfso.FileExists(path) then
         objfso.DeleteFile(path)   
	     set objfso = nothing
	    strstring = strstring &"<li>The file is deleted successfully from hard disk!"
                else
                       response.write "<script language='Javascript'>alert('The file have not found from hard disk ,it maybe deleted!')</script>"	      
                end if 
else
strstring = "<li>Parameter is wrong!"
end if
set rs=nothing
conn.close
set conn = nothing
response.write strstring
response.write "<li>After two seconds system is auto back…"
%>
<meta http-equiv="refresh" content='2;url=viewpro.asp'>
<%end if%>
</body>
</html>
