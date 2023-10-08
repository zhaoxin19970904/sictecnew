<!--#include file="conn.asp"-->
<%
if request("action")="login" then
admin_name=request("admin_name")
admin_pass=request("admin_pass")
if InStr(admin_name,"'") > 0 or InStr(admin_pass,"'") >0 then
%>
<script language="">
  alert("The characters are wrong!")
  location.href="<%=request.serverVariables("Http_REFERER")%>"
</script>
<%
else

set rs=server.createobject("adodb.recordset")
sql="select * from admin where admin_name='"&admin_name&"' and admin_pass='"&admin_pass&"'"
rs.open sql,conn,3,3
    if rs.eof then
%>
<script language="">
  alert("Your password or ID is wrong!")
  location.href="<%=request.serverVariables("Http_REFERER")%>"
</script>
<%
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