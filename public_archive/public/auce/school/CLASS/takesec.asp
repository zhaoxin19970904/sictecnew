<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
curclassid=trim(session("classid"))
curmonitor1=trim(request.form("monitor1"))
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
if userid<>monitor  then
   response.redirect "../error.asp?info=�Բ���������������Ա��Ȩ���룡"
end if  

if curmonitor1<> "" then
   sql="select * from userjoinclassinfo where classid='"&curclassid&"' and userid='"&curmonitor1&"'"
 
   rs.open SQL,schooldb
   if rs.eof then
      response.redirect "../error.asp?info=�Բ�������Ҫ�ӵĸ�����Ա���Ǳ����Ա��"
   end if
   rs.close 
end if
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb,1,3
if not rs.eof then
   rs("monitor1")=curmonitor1
   rs.update
end if
rs.close
response.redirect "../success.asp?info= �޸ĸ�����Ա�ɹ��ˣ�"  
%>