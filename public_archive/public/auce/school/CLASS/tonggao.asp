<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
curclassid=trim(session("classid"))
tonggao=trim(request.form("tonggao"))
if userid="" then
    response.redirect "../error.asp?info=�Բ������Ѿ������ˣ����������룡"
end if
set rs = createobject("ADODB.recordset")
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
   monitor=rs("monitor")
   monitor1=rs("monitor1")
end if
rs.close
if userid<>monitor and userid<>monitor1 then
   response.redirect "../error.asp?info=�Բ��������Ǳ���೤��Ȩ���룡"
end if
if len(tonggao)>300 then
   response.redirect "../error.asp?info=�Բ���ͨ���������ܳ������ٸ���"
end if  
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb,1,3
if not rs.eof then
   rs("tonggao")=tonggao
   rs.update
end if
rs.close
 response.redirect "../success.asp?info= �޸�ͨ��ɹ��ˣ�"     
%> 