<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
curclassid=trim(session("classid"))
curmonitor=trim(request.form("monitor"))
if userid="" then
    response.redirect "../error.asp?info=�Բ������Ѿ������ˣ����������룡"
end if
set rs = createobject("ADODB.recordset")
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
   monitor=rs("monitor")
end if
rs.close
if userid<>monitor then
   response.redirect "../error.asp?info=�Բ��������Ǳ������೤��Ȩ���룡"
end if  
if curmonitor="" then
   response.redirect "../error.asp?info=�Բ������εİ೤��������Ϊ�գ�"
end if 
sql="select * from userjoinclassinfo where classid='"&curclassid&"' and userid='"&curmonitor&"'"
rs.open SQL,schooldb
if rs.eof then
   response.redirect "../error.asp?info=�Բ�������Ҫ�������°೤���Ǳ����Ա��"
end if
rs.close 
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb,1,3
if not rs.eof then
   rs("monitor")=curmonitor
   rs.update
end if
rs.close
response.redirect "../success.asp?info= ����ְ�ɹ��ˣ�"

%> 