 <!--#include file="connect.asp"-->
<!--#include file="md5.asp" -->
<%
dim errmess
if request.Form("addadmin")<>"" then
call addadminuser()
end if
if request.Form("b1")<>"" then
call changepass()
end if
sub addadminuser()
dim user,psw,psw2,rs,sqltext
user=trim(replace(request.form("username"),"'",""))
psw=trim(replace(request.form("password1"),"'",""))
psw2=trim(replace(request.form("qqq"),"'",""))
set rs=conn.execute("select user from userinfo where user='"&user&"'")
if rs.eof or rs.bof then
errmess="<li>高级用户必须在注册用户内.增加管理员失败!"
exit sub
else
conn.execute("update userinfo set jibie='管理员' where user='"&user&"'")
end if
set rs=server.createobject("adodb.recordset")
sqltext="select * from admin"
rs.open sqltext,conn,1,3
if user="" or psw="" then
errmess="<li>用户名密码都不能为空!增加管理员失败!"
exit sub
end if
if psw<>psw2 then 
errmess="<li>两次密码输入不一样!增加管理员失败!"
exit sub
end if
rs.addnew

rs("username")=user
rs("password")=MD5(psw)
rs.update
errmess="<li>新管理员添加成功!"
end sub 

sub changepass() '更改高级管理员密码
dim username,password,password2,password3,rs,sql
if  request.form("username")<>""  then
username=trim(replace(request("username"),"'",""))
password=trim(replace(request("password"),"'",""))
password2=trim(replace(request("password2"),"'",""))
password3=trim(replace(request("password3"),"'",""))
if  password2="" or password3="" then
errmess="<li>新密码不能为空!修改失败!"
exit sub
end if
If password2 <>password3  Then
errmess="<li>两次输入的密码不一致!修改失败!"
exit sub
End If
set rs=server.createobject("adodb.recordset")
Sql="Select password,username from admin where username='"&username&"'"
Rs.Open Sql,Conn,1,3
if not rs.eof then
If Rs("password")<>MD5(password) Then
errmess="<li>您输入的密码不正确!修改失败!"
exit sub
End If
elseif  rs.eof  then
errmess="<li>您不是管理员不能修改!修改失败!"
exit sub
end if

Rs("password")=MD5(password2)
rs.update
end if
errmess="密码修改成功!请记新密码!"
end sub
%>
<html>
<head>
<title>论坛管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<div align="center"><br>
  <table width="450" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
    <%
call listadminuser()
sub listadminuser()
dim rs,sqltext,i
set rs=conn.execute("select username,id from admin")
do while not rs.eof
%>
    <tr bgcolor="#FFFFFF"> 
      <td width="198"> <%= rs("username")%> </td>
      <td width="167"><font color="#000000">高级管理员</font> </td>
      <td width="71"> 
        <div align="center"><a href=admin_cz.asp?id=<%=rs("id")%>&czlb=deladm title="删除此产品类别!(请不要轻易删除类别)" onclick="{if(confirm('确实要将该主论坛删除吗，确定删除吗?')){return true;}return false;}">删除 
          </a><a href="add_admin.asp?user=<%=rs("username")%>">修改</a></div>
      </td>
    </tr>
    <% 
rs.movenext
loop
end sub 
%>
  </table>
  
</div>
<form name="form1" method="post" action="add_admin.asp">
  <table width="450" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
    <tr><td bgcolor="#FFFFFF" colspan="2"></td></tr>
	<tr bgcolor="#92b9fb"> 
      <td height="27" colspan="2" background="backimg/bg1.gif"> 
        <div align="center"><font size="4"><b><font color="#FF0000">高级管理员添加</font></b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="62" bgcolor="#FFFFFF">用户名</td>
      <td width="367" bgcolor="#FFFFFF" align="right"> 
        <input name="username" type="text"  size="30" maxlength="30">
      </td>
    </tr>
    <tr> 
      <td width="62" bgcolor="#FFFFFF">密码</td>
      <td width="367" bgcolor="#FFFFFF" align="right"> 
        <input name="password1" type="password"  size="30" maxlength="30">
      </td>
    </tr>
    <tr> 
      <td width="62" bgcolor="#FFFFFF">重复密码</td>
      <td width="367" bgcolor="#FFFFFF" align="right"> 
        <input type="password" name="qqq" size="30" maxlength="30">
      </td>
    </tr>
    <tr> 
      <td width="62" bgcolor="#FFFFFF">　</td>
      <td width="367" bgcolor="#FFFFFF" align="right"> 
        <div align="center"> 
          <input name="addadmin" type="submit" id="addadmin" value="添 加">
        </div>
      </td>
    </tr>
  </table>

  
</form>  

</body>
</html>

