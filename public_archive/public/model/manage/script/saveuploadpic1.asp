<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.Buffer=true 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����ϴ�ͼƬ</title>
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
    '�Զ������ļ���
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
        file.saveAs Server.mappath("photo1/"&filename)'��ͼƬ
        response.write "ͼƬ�ϴ��ɹ����ļ���Ϊ:"
		response.write "<font color=red>" & filename &"</font><br><br>"
		response.write "<a href='temp1.asp'><u>ȷ��</u></a>&nbsp;&nbsp;"
    else
        response.write "ͼƬ�ϴ�ʧ�ܣ�ͼƬ�ļ���ʽ����!<br>"
        response.write "<a href='javascript:history.go(-1)'><u>���������ϴ�</u></a>"
    end if
    set file=nothing
else
    response.write "ͼƬ�ϴ�ʧ�ܣ��ļ�����Ϊ�գ�<br>"
    response.write "<a href='javascript:history.go(-1)'><u>���������ϴ�</u></a>"
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
