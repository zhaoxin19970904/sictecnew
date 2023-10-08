<!--#include file="CONN.ASP" -->
<%
dim rs,sqltext,czlb,user,tempout1,id
user=request.cookies("renwen")("user")
czlb=request.QueryString("czlb")
id=request.QueryString("id")
if czlb="" then
response.Redirect("error.asp?founderr=link")
end if
if user="" or request.cookies("renwen")("passedok")<>"ofdkjduy" then
errmess="<li>对不起!我们发现你现在还没登录!所以无法完成你的操作请求!请<a href=Login.asp>登录</a>"
response.cookies("rewin")("errmess")=errmess
response.Redirect("error.asp?founderr=mess")
end if

select case(czlb)
case "delmess"
if request.Form("pdel")="" or user="" then
errmess="<li>产生错误!<lb>可能你选择册除的短信息为空.<li>可能你没有登录.此操作权限为注册用户<a href=index.asp> 返回首页</a>"
response.cookies("rewin")("errmess")=errmess
response.Redirect("error.asp?founderr=mess")
end if
set rs=server.createobject("adodb.recordset")
sqltext="delete from mess where id in("&request.Form("pdel")&") and fd='"&user&"'"
rs.open sqltext,conn,1,3
response.Redirect("usermanager.asp?p=usersms")

case "addfriend" 
dim fuser,i,addok
addok=0
fuser=split(request.Form("friend"),",")
for i=0 to ubound(fuser)
set rs=conn.execute("select user from userinfo where user='"&trim(fuser(i))&"'")
if (not rs.eof) and (not rs.bof) then
    if conn.execute("select fd from fd where fd='"&trim(fuser(i))&"' and me='"&user&"'").eof then
	conn.execute("insert into fd ([fd],[me],[date]) values ('"&trim(fuser(i))&"','"&user&"',date())")
    addok=addok+1
	end if
else
errmess=fuser(i)&","&errmess
end if
next
if errmess<>"" then
tempout1=tempout1&"<li>"&errmess&" 用户没有被找到,看是否填写有误."
end if
if addok>0 then
tempout1=tempout1&"<li><b>"&addok&"</b> 用户被你成功的加为好友."
end if
response.cookies("rewin")("errmess")=(tempout1&"<li><li><a href=usermanager.asp?p=editfriend>返回好友列表</a><li><li><a href=index.asp>返回首页</a>")
response.Redirect("error.asp?founderr=mess")
case "delfriend"
conn.execute("delete from fd where id="&request.QueryString("id")&" and me='"&user&"'")
response.Redirect("usermanager.asp?p=editfriend")
case "addf"
set rs=conn.execute("select user,tzid from userf where id="&id)
if rs.eof or rs.bof then
conn.execute("insert into userf ([user],[tzid],[time]) values ('"&user&"',"&id&",now())")
end if
response.cookies("rewin")("errmess")=("<li>恭喜!你的帖子,已成功加入收藏夹.<li><li><a href=index.asp>返回首页</a><li><a href=type.asp?id="&id&">返回收藏的帖子</a>")
response.Redirect("error.asp?founderr=mess")
case "deluserf"
conn.execute("delete from userf where tzid="&request.QueryString("id")&" and user='"&user&"'")
response.Redirect("usermanager.asp?p=userf")
case "delfile"
set rs=conn.execute("select * from upfile where id="&id&" and user='"&user&"'")
if not rs.eof then
dim whichfile,fs,thisfile
whichfile=server.mappath(rs("pic"))
Set fs = CreateObject("Scripting.FileSystemObject")
Set thisfile = fs.GetFile(whichfile)
thisfile.Delete true
set fs=nothing
conn.execute("delete from upfile where id="&request.QueryString("id")&" and user='"&user&"'")
end if
response.Redirect("usermanager.asp?p=userpic")
end select
%>