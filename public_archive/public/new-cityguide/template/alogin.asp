
<%
if request("action")="login" then
admin_name=request("admin_name")
admin_pass=request("admin_pass")
if InStr(admin_name,"'")> 0 or InStr(admin_pass,"'")>0 then
   response.write "Invalid character"
   response.end
end if
%>
<%
'set rs=server.createobject("adodb.recordset")
'sql="select * from admin where admin_name='"&admin_name&"' and admin_pass='"&admin_pass&"'"
'rs.open sql,conn,3,3 ' if rs.eof then

if admin_name<>"chinacity" or admin_pass<>"sicteccity" then
        response.write "<center>Username and password not match."
else
        session("user_name")="sictec"
        response.redirect "index.asp"
end if
end if
'rs.close
'set rs=nothing
'conn.close
'set conn=nothing
'end if
%>