<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
curclassid=trim(session("classid"))
curmonitor=trim(request.form("monitor"))
if userid="" then
    response.redirect "../error.asp?info=对不起，您已经掉线了，请重新申请！"
end if
set rs = createobject("ADODB.recordset")
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
   monitor=rs("monitor")
end if
rs.close
if userid<>monitor then
   response.redirect "../error.asp?info=对不起，您不是本班正班长无权进入！"
end if  
if curmonitor="" then
   response.redirect "../error.asp?info=对不起，新任的班长姓名不能为空！"
end if 
sql="select * from userjoinclassinfo where classid='"&curclassid&"' and userid='"&curmonitor&"'"
rs.open SQL,schooldb
if rs.eof then
   response.redirect "../error.asp?info=对不起，您所要任命的新班长不是本班成员！"
end if
rs.close 
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb,1,3
if not rs.eof then
   rs("monitor")=curmonitor
   rs.update
end if
rs.close
response.redirect "../success.asp?info= 您辞职成功了！"

%> 