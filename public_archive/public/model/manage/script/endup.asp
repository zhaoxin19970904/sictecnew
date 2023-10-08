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
sign = request.querystring("sign")
deal=request.querystring("deal")

if deal="big" then
set rs=server.createobject("adodb.recordset")
strsql = "select photo2 from modeldb where sign="&sign&""
rs.open strsql,conn,1,1
url=rs(0)
rs.close
path=server.mappath("photo2/"&url)

strsql = "delete * from modeldb where sign="&sign&""
conn.execute strsql,intno
if intno <> 0 then
   strstring = "<li>该记录已成功从数据库中删除!"
   set objfso = server.createobject("scripting.FileSystemObject")
     if objfso.FileExists(path) then
         objfso.DeleteFile(path)   
	     set objfso = nothing
	    strstring = strstring &"<li>文件从硬盘清除成功!"
                else
                       response.write "<script language='Javascript'>alert('指定的文件在硬盘上并未找到,可能已被清除\n请核实!')</script>"	      
                end if 
else
strstring = "<li>非法参数!"
end if
set rs=nothing
conn.close
set conn = nothing
response.write strstring
response.write "<li>系统将会在2秒内自动返回…"
%>
<meta http-equiv="refresh" content='2;url=view.asp'>
<%
elseif deal="small" then
set rs=server.createobject("adodb.recordset")
strsql = "select photo2,photo1 from modeldb where sign="&sign&""
rs.open strsql,conn,1,1
url=rs(0)
url1=rs(1)
rs.close
path=server.mappath("photo2/"&url)
path1=server.mappath("photo1/"&url1)

strsql = "delete * from modeldb where sign="&sign&""
conn.execute strsql,intno
if intno <> 0 then
   strstring = "<li>该记录已成功从数据库中删除!"
   set objfso = server.createobject("scripting.FileSystemObject")
     if objfso.FileExists(path) then
         objfso.DeleteFile(path)
		 objfso.DeleteFile(path1)    
	     set objfso = nothing
	    strstring = strstring &"<li>文件从硬盘清除成功!"
                else
                       response.write "<script language='Javascript'>alert('指定的文件在硬盘上并未找到,可能已被清除\n请核实!')</script>"	      
                end if 
else
strstring = "<li>非法参数!"
end if
set rs=nothing
conn.close
set conn = nothing
response.write strstring
response.write "<li>系统将会在2秒内自动返回…"
%>
<meta http-equiv="refresh" content='2;url=viewmodel.asp'>

<%
end if
%>

</body>
</html>
