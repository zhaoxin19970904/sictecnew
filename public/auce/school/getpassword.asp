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
   response.redirect "error.asp?info=对不起，此用户不存在！！！"   
end if 
rs.close
if curpassask<>passask or curanwserpass<>anwserpass then
   response.redirect "error.asp?info=对不起，您所填写的密码提示问题或回答不正确！！！"   
else
   response.redirect "success.asp?info=您的密码是<font color=red>"&passwd&"</font>,请您记住您的密码！！！"
end if   
%>