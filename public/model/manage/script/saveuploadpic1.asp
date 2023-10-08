<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.Buffer=true 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>保存上传图片</title>
<link href=style.css rel=STYLESHEET type=text/css>
</head>
<body>
<%
if session("admin_name")="" then Response.Redirect("../index.asp")
%>
<!--#include FILE="upload_5xsoft.inc"--> 
<!--#include FILE="conn.asp" --> 

<%
set rs=server.CreateObject("adodb.recordset")
id=session("id")
sql="select photo1 from modeldb where id="&id&""
rs.open sql,conn,1,1
session("photo1")=rs("photo1")
rs.close




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
        file.saveAs Server.mappath("photo1/"&filename)'大图片
        response.write "图片上传成功！文件名为:"
		response.write "<font color=red>" & filename &"</font><br><br>"
		response.write "<a href='temp1.asp'><u>确定</u></a>&nbsp;&nbsp;"
    else
        response.write "图片上传失败！图片文件格式不对!<br>"
        response.write "<a href='javascript:history.go(-1)'><u>返回重新上传</u></a>"
    end if
    set file=nothing
else
    response.write "图片上传失败！文件内容为空！<br>"
    response.write "<a href='javascript:history.go(-1)'><u>返回重新上传</u></a>"
end if
set upload=nothing

sql="select photo1,update from modeldb where id="&id&""
rs.open sql,conn,1,3
rs("photo1")=filename
rs("update")=now
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</body> 
</html>
