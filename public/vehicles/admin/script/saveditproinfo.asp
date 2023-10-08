<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=楷体_GB2312">
<title>添加信息</title>
</head>
<body>
<%
 id=request.Form("product_id")
 p_type=ltrim(rtrim(request.form("type")))
 model=ltrim(rtrim(request.form("model")))
 price=request.Form("price")
 Note=request.Form("Note")
 
 txt_long=len(note)
 for i=1 to txt_long
	  note=replace(note,vbCrLf,"<br>")
next
 
 ifcommended=request.Form("ifcommended")
 set rs=server.createobject("adodb.recordset")
 sql="select * from products where id="&id&""
 rs.open sql,conn,1,3
 rs("type")=p_type
 rs("modelno")=model
 rs("price")=price
 if trim(ifcommended)="yes" then
   rs("ifcommended")=true
 else
   rs("ifcommended")=false
 end if
  rs("introduction")=Note
  rs("up_date")=now()
  rs("up_admin")=session("admin_name")
 rs.update
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 'response.write "The product information have saved!<br><br>"
 'response.write "<a href='upload_pic.asp'>Continue up new</a>&nbsp;&nbsp;&nbsp;&nbsp;"
 'response.write "<a href='viewpro.asp'>View the product</a>"
%>
<script language="JavaScript">
  alert("Successfully!")
  location.href="viewpro.asp"
</script>
</body>
</html>
