<!--#include file="admin_conn.asp" -->
<!--#include file="md5.asp" -->
<%
if request.Form("username")<>"" then
call checklogin()
end if
if request.Form("adminname")<>"" then
call adminlogin()
end if

sub checklogin()
dim user,upwd,rs,usercookies,gourl,userip
userip=request.ServerVariables("REMOTE_ADDR")
user=trim(replace(request.Form("username"),"'",""))
upwd=trim(replace(request.Form("pwd"),"'",""))
usercookies=cint(request.form("cookies"))

set rs=conn.execute("select user,password,id from userinfo where user='"&user&"'")
if not rs.eof then 
     if rs("password")<>MD5(upwd) then
	 errmess=errmess&"<li>你的用户名或者密码不正确!"
	exit sub
	end if
select case usercookies
case 0
	Response.Cookies("renwen")("usercookies") = usercookies
case 1
   	Response.Cookies("renwen").Expires=Date+1
	Response.Cookies("renwen")("usercookies") = usercookies
case 2
	Response.Cookies("renwen").Expires=Date+31
	Response.Cookies("renwen")("usercookies") = usercookies
case 3
	Response.Cookies("renwen").Expires=Date+365
	Response.Cookies("renwen")("usercookies") = usercookies
case other
	Response.Cookies("renwen").Expires=date+0.01
	Response.Cookies("renwen")("usercookies") = usercookies
end select
response.cookies("renwen")("user")=rs("user")
response.cookies("renwen")("passedok")="ofdkjduy"
conn.execute("update userinfo set userlog=userlog+1,userjf=userjf+1,lastlogin='"&now()&"',lastlogIP='"&userip&"' where user='"&user&"'")
	errmess=errmess&"<li><b>你已经成功登录.</b><li><a href=index.asp>返回首页</a>"
	rs.close
	else
	errmess=errmess&"<li>你的用户名或者密码不正确!"
    rs.close
  end if
end sub

sub adminlogin()
dim username,password,sql,rs
username=trim(replace(request("adminname"),"'",""))
password=trim(replace(request("password"),"'",""))
if  password<> "" and username<>"" then
sql="select username,password from admin where username='"&username&"' and password='"&MD5(password)&"'"
set rs=server.createobject("adodb.recordset")
set rs=conn.execute(sql)
if not rs.EOF then
session("user")=rs("username")
session("adpassdf")="passed"
response.redirect "admin_index1.asp"
else 
errmess=errmess&"管理密码不正确,请重新输入！"
end if
end if
end sub
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="database/Style.css">
<title>管理登陆</title>
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<form method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>" name="adlogin">
  <p align="center"><%=errmess%></p>
  <%if (not userislogin) or (not master) then 
  if userislogin then
  response.Write("<p align=center>你目前不是管理员不具有管理权限!</p>")
  end if
  %>
  <table width="327" align="center" cellpadding=3 cellspacing=1 bgcolor="#92b9fb" class=fs>
    <tr align="center"> 
      <td height="27" colspan="2" background="backimg/bg1.gif"> <font color="#FFFFFF"><strong>确定管理身份</strong></font></td>
    </tr>
    <tr> 
      <td width="97" bgcolor="#F2F2F2"> <font color="#000033">用户名:</font></td>
      <td width="213" bgcolor="#FFFFFF"> <font color="#000033"> 
        <input class=inputline name=username size=15>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor="#F2F2F2"> <font color="#000033">密码:</font></td>
      <td bgcolor="#FFFFFF"> <font color="#000033"> 
        <input class=inputline name=pwd size=15 
            type=password>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor="#F2F2F2"><font color="#000033">cookies保存</font></td>
      <td bgcolor="#FFFFFF"> <font color="#000033"> 
        <select name="select" id="select2">
          <option value="0">不保存</option>
          <option value="1">一天</option>
          <option value="2">一个月</option>
          <option value="3">一年</option>
        </select>
        </font></td>
    </tr>
    <tr align="center"> 
      <td colspan=2 bgcolor="#FFFFFF"> 
        <input name=form1 type=submit value=" 登 录 "> &nbsp;&nbsp; 
        <input class=main name=reset2 type=reset value=" 重 置 "> </td>
    </tr>
  </table>
  <%else%>
  <table width="321" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
    <tr align="center"> 
      <td height="28" colspan="2" background="backimg/bg1.gif"><font color="#FFFFFF"><b>管理登陆</b></font> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="83">管理员名称：</td>
      <td width="223"> <input type=text name=adminname size=20 maxlength=20 class=link> 
        <font color="#FF3333">*</font> 字母或数字</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>管理员密码：</td>
      <td> <input type=password name=password size=20 maxlength=20 class=link> 
        <font color="#FF3333">*</font> 字母或数字</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;</td>
      <td> <input type="submit" value="进入!" name=adlogin></td>
    </tr>
  </table>
  <%end if%>
</form>
</html>
