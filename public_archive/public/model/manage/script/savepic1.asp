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
        file.saveAs Server.mappath("photo1/"&filename)'СͼƬ
        response.write "ͼƬ�ϴ��ɹ����ļ���Ϊ:"
		response.write "<font color=red>" & filename &"</font><br><br>"
		response.write "<a href='modelinfo.asp'><u>�����ϴ�ģ�͵�������Ϣ</u></a>&nbsp;&nbsp;"
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

'mdbfile=server.MapPath("data/upfile.mdb")
 'set conn=server.CreateObject("adodb.connection")
'conn.open "driver={microsoft access driver (*.mdb)};dbq="&mdbfile

max_sign=session("max_sign")
set rs=server.CreateObject("adodb.recordset")
sql="select * from modeldb where sign="&max_sign&""
rs.open sql,conn,1,3
rs("photo1")=filename
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<a href="endup.asp?deal=small&sign=<%=max_sign%>"><u>�����ϴ�</u></a>
</body> 
</html>