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
sql="select max(sign) from modeldb"
rs.open sql,conn,1,1
session("max_sign")=rs(0)+1
max_sign=session("max_sign")
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
        file.saveAs Server.mappath("photo2/"&filename)'大图片
        response.write "图片上传成功！文件名为:"
		response.write "<font color=red>" & filename &"</font><br><br>"
		response.write "<a href='upload_pic1.asp'><u>继续上传小图片</u></a>&nbsp;&nbsp;"
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

'mdbfile=server.MapPath("data/upfile.mdb")
 'set conn=server.CreateObject("adodb.connection")
'conn.open "driver={microsoft access driver (*.mdb)};dbq="&mdbfile

sql="select * from modeldb"
rs.open sql,conn,1,3
rs.addnew
rs("photo2")=filename
rs("update")=now
rs("sign")=max_sign
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<a href="endup.asp?deal=big&sign=<%=max_sign%>"><u>结束上传</u></a>
</body> 
</html>
