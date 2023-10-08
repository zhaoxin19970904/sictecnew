<!--#include file="CONN.ASP" -->
<!--#include file="md5.asp" -->
<%
if request.Form("username")<>"" then
call checklogin()
end if
sub checklogin()
dim user,upwd,rs,usercookies,gourl,userip
userip=request.ServerVariables("REMOTE_ADDR")
gourl=request.Form("thispagename")
user=trim(replace(request.Form("username"),"'",""))
upwd=trim(replace(request.Form("pwd"),"'",""))
usercookies=cint(request.form("cookies"))
if gourl="" or left(gourl,9)="login.asp" then
gourl="index.asp"
end if

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
	founderr=true
	errmess=errmess&"<li><b>你已经成功登录.</b> 三秒自动返回首页.<li><li><a href=index.asp>返回首页</a>"
	errmess=errmess&"<script>window.tm = setInterval(""location.href='index.asp'"", 3000)</script>"
	rs.close
	else
	errmess=errmess&"<li>你的用户名或者密码不正确!"
    rs.close
  end if
end sub
if request.QueryString("userlogout")="loginout" then
dim userip
userip=request.ServerVariables("REMOTE_ADDR")
conn.execute=("update userinfo set logout=logout+1,lastlogIP='"&userip&"' where user='"&request.cookies("renwen")("user")&"'")
Response.Cookies("renwen").Expires=Date+0.001
response.cookies("renwen")("user")=""
response.cookies("renwen")("passedok")=""
userislogin=false
	founderr=true
	errmess=errmess&"<li><b>你已经成功退出登录.</b> 三秒自动返回首页.<li><li><a href=index.asp>返回首页</a>"
	errmess=errmess&"<script>window.tm = setInterval(""location.href='index.asp'"", 3000)</script>"
end if
%>
<!--#include file="mymem.asp" -->    
<FORM  action=Login.asp  method=post name="Login">
<%
if founderr=true then
response.cookies("rewin")("errmess")=errmess
response.Redirect("error.asp?founderr=mess")
end if
%>
  <table cellpadding=3 cellspacing=1 class=fs height=219 width="898" align="center" bgcolor="#92b9fb">
    <tr><td bgcolor="#FFFFFF" colspan=2>
	<%if errmess<>"" then response.Write(errmess) end if%></td></tr>
	<tr bgcolor="#006699"> 
      <td height="27" colspan="2" background="backimg/bg1.gif"> 
        <div align="center"><font color="#FFFFFF"><strong>用户登录</strong></font></div>
      </td>
    </tr>
    <tr bgcolor="#E0E4FE"> 
      <td width="162" height="24" bgcolor="#F2F2F2"> 
        <div align="center"><font color="#000033">用户名:</font></div>
      </td>
      <td width="384" height="24" bgcolor="#FFFFFF"> <font color="#000033"> 
        <input class=inputline name=username size=15>
        <font size="2"><a href="regedit.asp">没有注册</a></font></font></td>
    </tr>
    <tr bgcolor="#E0E4FE"> 
      <td width="162" height="9" bgcolor="#F2F2F2"> 
        <div align="center"><font color="#000033">密&nbsp;&nbsp;码:</font></div>
      </td>
      <td width=384 height="9" bgcolor="#FFFFFF"> <font color="#000033"> 
        <input class=inputline name=pwd size=15 
            type=password>
        <font size="2"><a href="lostpassport.asp">忘记密码</a>?</font></font></td>
    </tr>
    <tr bgcolor="#E0E4FE"> 
      <td width="162" height="96" bgcolor="#F2F2F2"><font color="#000033">请选择cookies在本地浏览器中保存的时间（在网吧上网的用户请用默认选择）。</font></td>
      <td width="384" height="96" bgcolor="#FFFFFF"> <font color="#000033"> 
        <input type="radio" name="cookies" value="0" checked>
        不保存 ，关闭浏览器就失效<br>
        <input type="radio" name="cookies" value="1">
        一天<br>
        <input type="radio" name="cookies" value="2">
        一个月<br>
        <input type="radio" name="cookies" value="3">
        一年 </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="33" colspan=2> 
        <div align="left"></div>
        <table width="188" height="26" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td> 
           
                <input name=form1 type=submit value=" 登 录 ">
              </td>
            <td> 
                <input class=main name=reset2 type=reset value=" 重 置 ">
              </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</FORM >
<!--#include file="end.asp" -->
</BODY>
</HTML>
