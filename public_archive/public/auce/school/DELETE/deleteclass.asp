<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
if session("pass")<>"ok" then
   response.redirect "../error.asp?info=您无权进入！"
end if 
curclassid=trim(request("classid"))
set rs = createobject("ADODB.recordset")
sql="delete from userjoinclassinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb,1,3
sql="delete from message where classid='"&curclassid&"'"
rs.open SQL,schooldb,1,3
'sql="delete from upload where classid='"&curclassid&"'"
'rs.open SQL,schooldb,1,3
sql="delete from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb,1,3
response.redirect "../success.asp?info=注销该班级成功了！"
%>   


