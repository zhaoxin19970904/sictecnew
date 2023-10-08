<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.Buffer=true 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=楷体_gb2312">
<title>保存上传图片</title>
<link href=style.css rel=STYLESHEET type=text/css>
</head>
<body>
<%
if session("admin_name")="" then Response.Redirect("../ ../index.asp")
%>
<!--#include FILE="upload_5xsoft.inc"--> 
<!--#include FILE="conn.asp" --> 
<%
set rs=server.CreateObject("adodb.recordset")
set upload=new upload_5xsoft
set file=upload.file("file1")
if file.fileSize>0 then
    '自动生成文件名
    filename=date()
    filename=filename&time()
    filename=replace(filename,"-","")
    filename=replace(filename,":","")
    filename=replace(filename," ","")
    filename=filename+"."
    filenameend=file.filename
    filenameend=split(filenameend,".")
    if Ucase(filenameend(1))="GIF" or Ucase(filenameend(1))="JPG" then
        filename=filename&filenameend(1)
        file.saveAs Server.mappath("../../proimages/"&filename)
        response.write "Up file successfully！<br>The file name is:"
		response.write "<font color=red>" & filename &"</font><br><br>"
		response.write "<A href='proinfo.asp'>Up the product's other information.</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    else
        response.write "Up photo is defeated ！The format of this photo is wrong!<br>"
        response.write "<a href='javascript:history.go(-1)'><u>Back</u></a>"
    end if
    set file=nothing
else
    response.write "Up photo is defeated ！The file content is empty！<br>"
    response.write "<a href='javascript:history.go(-1)'><u>Back</u></a>"
end if
set upload=nothing
sql="select photo,up_date,up_admin from products"
rs.open sql,conn,1,3
rs.addnew
rs("photo")=filename
rs("up_date")=now
rs("up_admin")=session("admin_name")
rs.update
session("filename")=filename
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<a href="endup.asp?filename=<%=filename%>"><u>End up product</u></a>
</body> 
</html>
