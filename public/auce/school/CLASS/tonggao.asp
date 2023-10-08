<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
curclassid=trim(session("classid"))
tonggao=trim(request.form("tonggao"))
if userid="" then
    response.redirect "../error.asp?info=对不起，您已经掉线了，请重新申请！"
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
   response.redirect "../error.asp?info=对不起，您不是本班班长无权进入！"
end if
if len(tonggao)>300 then
   response.redirect "../error.asp?info=对不起，通告字数不能超过三百个！"
end if  
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb,1,3
if not rs.eof then
   rs("tonggao")=tonggao
   rs.update
end if
rs.close
 response.redirect "../success.asp?info= 修改通告成功了！"     
%> 