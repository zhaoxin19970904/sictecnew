<!--#include file="admin_conn.asp" -->
<%
dim czlb,id,rs,sqltext
czlb=request.QueryString("czlb")
id=request.QueryString("id")
select case(czlb)
case"dellb"
conn.execute("delete from ltname where id="&id)
response.Redirect("admin_index.asp")
case"delzlt"
conn.execute("delete from ltlb where id="&id)
response.Redirect("admin_index.asp")
case "offzlt"
set rs=conn.execute("select ifoff from ltlb where id="&id)
if rs("ifoff")=0 then 
conn.execute("update ltlb set ifoff=1 where id="&id)
else
conn.execute("update ltlb set ifoff=0 where id="&id)
end if
response.Redirect("admin_index.asp")
case "deluser"
conn.execute("delete from userinfo where id="&id)
response.Redirect("listalluser.asp")
case "quitadmin"
session.abandon()
Response.Redirect "index.asp"
case "offlt"
set rs=conn.execute("select offlt from admin_copyc where id="&id)
if rs("offlt")=0 then 
conn.execute("update admin_copyc set offlt=1 where id="&id)
else
conn.execute("update admin_copyc set offlt=0 where id="&id)
end if
response.Redirect("admin_copyc.asp")
case "deladm"
set rs=server.createobject("adodb.recordset")
rs.open "select username from admin",conn,1,1
if rs.recordcount<=1 then
response.Write("高级管理员不能少于一个!")
response.End()
else
conn.execute("delete from admin where id="&id)
response.Redirect("add_admin.asp")
end if
end select
%>
