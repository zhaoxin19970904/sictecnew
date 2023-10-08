<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
if session("pass")<>"ok" then
   response.redirect "../error.asp?info=您无权进入！"
end if 
schoolid=trim(request("schoolid"))
set rs = createobject("ADODB.recordset")
sql="delete from userjoinclassinfo where classid like '"&schoolid&"%'"
rs.open SQL,schooldb,1,3
sql="delete from message where classid like '"&schoolid&"%'"
rs.open SQL,schooldb,1,3
sql="delete from classinfo where schoolid = '"&schoolid&"'"
rs.open SQL,schooldb,1,3
sql="delete from schoolinfo where schoolid = '"&schoolid&"'"
rs.open SQL,schooldb,1,3
response.redirect "../success.asp?info=注销该学校成功了！"
%>   
