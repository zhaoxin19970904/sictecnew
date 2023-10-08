<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加信息</title>
</head>
<body>
<%
id=request.querystring("id")
 m_Type=request.form("m_Type")
 m_Name=request.form("m_Name")
 Manufacturer=request.form("Manufacturer")
 original=request.form("original")
 Rato=request.form("Rato")
 m_Size=request.form("m_Size")
 Org_size=request.form("Org-size")
 Material=request.form("Material")
 Price=request.form("Price")
 Miniquantity=request.form("Miniquantity")
 Leadtime=request.form("Leadtime")
 Packing=request.form("Packing")
 Sample=request.form("Sample")
 Note=request.form("Note")
 width_height=request.form("width_height")
 max_sign=session("max_sign")
 set rs=server.createobject("adodb.recordset")
 sql="select * from modeldb where id="&id&""
 rs.open sql,conn,1,3
 rs("Type")=m_Type
 rs("Name")=m_Name
 rs("Manufacturer")=Manufacturer
 rs("original")=original
 rs("Ratio")=Rato
 rs("Size")=m_Size
 rs("Org-size")=Org_size
 rs("Material")=Material
 rs("Price")=Price
 rs("Miniquantity")=Miniquantity
 rs("Leadtime")=Leadtime
 rs("packing")=packing
 rs("update")=now
 
 if trim(Sample)="有" then
 rs("Sample")=true
 else
 rs("Sample")=false
 end if
  rs("Notes")=Note
  rs("w_h")=width_height
 rs.update
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 response.write "模型修改后的信息已成功保存<br>"
 response.write "<a href='viewmodel.asp'>查看模型</a>"
%>
</body>
</html>
