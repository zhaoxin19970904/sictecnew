<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
curclassid=trim(session("classid"))
pername=trim(request("pername"))
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
sql="select * from userjoinclassinfo where userid='"&pername&"' and classid='"&curclassid&"'"

rs.open SQL,schooldb
if rs.eof then
   response.redirect "../error.asp?info=�Բ��𣬴��˱����Ͳ��Ǳ����Ա��"
end if
rs.close
sql="delete from userjoinclassinfo where userid='"&pername&"' and classid='"&curclassid&"'"
rs.open SQL,schooldb,1,3
response.redirect "../success.asp?info=�����ɹ�,"&pername&" �Ѿ�����������ð࣡"
%> 