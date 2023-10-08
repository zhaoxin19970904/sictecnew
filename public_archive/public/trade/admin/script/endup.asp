<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>UPLOAD END</title>
</head>
<body>
<%
On error resume next
filename=trim(request.QueryString("filename"))
id=trim(request.QueryString("id"))
if filename<>"" then
set rs=server.createobject("adodb.recordset")
path=server.mappath("prophotos/"&filename)
strsql = "delete * from products where photo='"&filename&"'"
conn.execute strsql,intno
if intno <> 0 then
   strstring = "<li>Record has been deleted successfully!"
   set objfso = server.createobject("scripting.FileSystemObject")
     if objfso.FileExists(path) then
         objfso.DeleteFile(path)   
	     set objfso = nothing
	    strstring = strstring &"<li>File has been removed from hard disk!"
                else
                       response.write "<script language='Javascript'>alert('File has not been found.\n Please check!')</script>"	      
                end if 
else
strstring = "<li>Invalid Parameter!"
end if
set rs=nothing
conn.close
set conn = nothing
response.write strstring
response.write "<li>Return back in 2 seconds…"
%>
<meta http-equiv="refresh" content='2;url=viewpro.asp'>
<%

else
set rs=conn.execute("select photo from products where id="&id&"")
path=server.mappath("prophotos/"&rs(0))
strsql = "delete * from products where id="&id&""
conn.execute strsql,intno
if intno <> 0 then
   strstring = "<li>Record has been removed from database!"
   set objfso = server.createobject("scripting.FileSystemObject")
     if objfso.FileExists(path) then
         objfso.DeleteFile(path)   
	     set objfso = nothing
	    strstring = strstring &"<li>Removed successfully!"
                else
                       response.write "<script language='Javascript'>alert('指定的文件在硬盘上并未找到,可能已被清除\n请核实!')</script>"	      
                end if 
else
strstring = "<li>Invalid Parameter!"
end if
set rs=nothing
conn.close
set conn = nothing
response.write strstring
response.write "<li>Return back in 2 seconds…"
%>
<meta http-equiv="refresh" content='2;url=viewpro.asp'>
<%end if%>
</body>
</html>
