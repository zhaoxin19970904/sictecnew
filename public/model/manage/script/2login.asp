<!--#include file="conn.asp"-->
<%
if request("action")="login" then
admin_name=request("admin_name")
admin_pass=request("admin_pass")
if InStr(admin_name,"'") > 0 or InStr(admin_pass,"'") >0 then
response.write "请不要使用非法字符"
else
%><%
set rs=server.createobject("adodb.recordset")
sql="select * from admin where admin_name='"&admin_name&"' and admin_pass='"&admin_pass&"'"
rs.open sql,conn,3,3
    if rs.eof then
        response.write "<center>用户名和密码不匹配"
    else
        session("admin_name")=rs("admin_name")
        response.redirect "../index.asp"
    end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
end if
%>