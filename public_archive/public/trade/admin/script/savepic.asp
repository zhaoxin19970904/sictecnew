<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.Buffer=true 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html">
<title>save picture</title>
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
set upload=new upload_5xsoft
set file=upload.file("file1")
if file.fileSize>0 then
    'Auto Generating File name
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
        file.saveAs Server.mappath("prophotos/"&filename)
        response.write "Success£¡<br>Filename:"
		response.write "<font color=red>" & filename &"</font><br><br>"
		response.write "<A href='proinfo.asp'>Upload More info</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    else
        response.write "Failure! Invalid format.<br>"
        response.write "<a href='javascript:history.go(-1)'><u>Back</u></a>"
    end if
    set file=nothing
else
    response.write "Failure! Null file.<br>"
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
<a href="endup.asp?filename=<%=filename%>"><u>Finished</u></a>
</body> 
</html>
