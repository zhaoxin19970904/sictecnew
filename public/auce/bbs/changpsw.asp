<!--#include file="CONN.ASP" -->
<!--#include file="md5.asp" -->
<%
if not userislogin then
errmess=errmess&error1
call founderror(errmess)
end if
if request.form("user")<>"" then
call changepsw()
end if
sub changepsw()
dim user,password,password2,password3,rs,sql
user=trim(replace(request.Form("user"),"'",""))
password=trim(replace(request.Form("password"),"'",""))
password2=trim(replace(request.Form("password2"),"'",""))
password3=trim(replace(request.Form("password3"),"'",""))
set rs=conn.execute("select password,user from userinfo where user='"&user&"'")
if rs.eof or rs.bof then
errmess=errmess&"用户不存在，请注意用户名是否有区分全角字符，或多余的空格"
exit sub
end if
If Rs("password")<> MD5(password)  Then
errmess=errmess&"您输入的旧密码不正确"
exit sub
End If
If password2 <>password3  Then
errmess=errmess&"您两次输入的密码不一致！"
exit sub
End If
if Len(Password2)<6 then
errmess=errmess&"密码位数不能少于6位."
exit sub
end if
conn.execute("update userinfo set [password]='"&MD5(password2)&"' where user='"&user&"'")
	founderr=true
    errmess=errmess&"<li><b>密码修改成功! </b> 请记住新密码!"
	errmess=errmess&"<li>请返回首页重登录.三秒自动返回首页.<li><li><a href=index.asp>返回首页</a>"
	errmess=errmess&"<script>window.tm = setInterval(""location.href='index.asp'"", 3000)</script>"

end sub
%>
<!--#include file="mymem.asp" -->
<!--#include file="mymeumu.asp" -->
<%
if founderr=true then 
call founderror(errmess)
end if
%>

  
<table width=898 border=0 align="center" cellpadding=4 cellspacing=1 bgcolor="#92b9fb">
  <form method="post" action="changpsw.asp" name="form1">
    <tr align="center"> 
      <td height="27" colspan="2" background="backimg/bg1.gif" class=link><b><font color="#FFFFFF">用户更改登录密码</font></b></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td colspan="2" class=link><%if errmess<>"" then response.write(errmess) end if%></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td width=223 class=link>用户名称：</td>
      <td> <input name=user type=text class=link id="user2" size=20 maxlength=16> 
        <font color="#FF3333">*</font> 字母或数字</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td width=223 class=link>旧 密 码：</td>
      <td width=539> <input type=password name=password size=20 maxlength=16 class=link> 
        <font color="#FF3333">*</font> 字母或数字</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td width=223 class=link>新 密 码：</td>
      <td width=539> <input type=password name=password2 size=20 maxlength=16 class=link> 
        <font color="#FF3333">*</font> 字母或数字</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td width=223 class=link>确认密码：</td>
      <td width=539> <input type=password name=password3 size=20 maxlength=16 class=link> 
        <font color="#FF3333">*</font> 字母或数字</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td colspan="2" class=link> <input type="submit" value="确定修改" name="b1">
        　 
        <input type="reset" name="Reset" value="重新填写"> </td>
    </tr>
	</form>
  </table>


<!--#include file="end.asp" -->
</body>
</html>