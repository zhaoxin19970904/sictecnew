<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(request.form("userid"))
curpassask=trim(request.form("passask"))
curanwserpass=trim(request.form("anwserpass"))
set rs = createobject("ADODB.recordset")
sql="select * from userinfo where userid='"&userid&"'"
rs.open SQL,schooldb
if not rs.eof then
   passask=rs("passask")
   anwserpass=rs("anwserpass")
   passwd=rs("userpassword")
else
   response.redirect "error.asp?info=�Բ��𣬴��û������ڣ�����"   
end if 
rs.close
if curpassask<>passask or curanwserpass<>anwserpass then
   response.redirect "error.asp?info=�Բ���������д��������ʾ�����ش���ȷ������"   
else
   response.redirect "success.asp?info=����������<font color=red>"&passwd&"</font>,������ס�������룡����"
end if   
%>