<!--#include file="CONN.ASP" -->
<!--#include file="mymem.asp" -->
<!--#include file="md5.asp" -->
<%
dim rs,action
action=request.Form("action")
if action="" then
action="step1"
end if
if action="step4"then
call changepassword()
end if
sub changepassword()
dim password1,password2,user
user=request.Form("user")
password1=Trim(replace(request.Form("password1"),"'",""))
password2=Trim(replace(request.Form("password2"),"'",""))
if password1<>password2 then
founderr=true
errmess=errmess&"<li>你两次输入的密码不一至,请重新输入"
exit sub
end if
if Len(password1)<6 then
errmess=errmess&"<li>密码不能少于6位字符,请重新输入"
exit sub
end if
if request.Form("user")="" then
errmess=errmess&"<li>用户不能为空,请重新输入"
exit sub
end if
set rs=conn.execute("select quesion,answer,user from userinfo where user='"&user&"'")
if rs.eof or rs.bof then
errmss=errmess&"<li>发现错误,输入用户名不存在.请重新输入."
exit sub
end if
if rs("answer")=MD5(trim(request.Form("answer"))) then
conn.execute("update userinfo set [password]='"&MD5(password2)&"' where user='"&user&"'")
errmess=errmess&"<li><b>密码已经修改成功!</b> 请重新登录,并记住你的新密码"
else
errmess=errmess&"<li>发现错误,你输入问题答案不正确.请重新输入."
end if
end sub
%>
<form name="form" method="post" action="lostpassport.asp">
  <table width="898" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
    <%
	select case action
	case "step1"
	%>
    <tr bgcolor=#FFFFFF> 
      <td height="26" colspan="2" background="backimg/bg1.gif"> <b><font color="#FFFFFF">会员找回密码</font></b> 
        第一步 (<strong>输入用户名</strong>)</td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=9>用户名: </td>
      <td><input name="user" type="text" id="user"> <input name="action" type="hidden" id="action" value="step2"> 
      </td>
    </tr>
    <%
	case "step2"
	set rs=conn.execute("select quesion,answer,user from userinfo where user='"&request.form("user")&"'")
	if (not rs.eof) and (not rs.bof) then
	%>
	    <tr bgcolor=#FFFFFF> 
      <td height="26" colspan="2" background="backimg/bg1.gif"> <b><font color="#FFFFFF">会员找回密码</font></b> 
        第二步 (<strong>输入问题答案</strong>)</td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=4>用户名:</td>
      <td><%=rs("user")%></td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=4>问 题:</td>
      <td><%=rs("quesion")%></td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=9>答 案: </td>
      <td> <input name=answer type=text id=answer size=30> <input name="user" type="hidden" id="action" value="<%=rs("user")%>"> 
        <input name="action" type="hidden" id="action" value="step3"> </td>
    </tr>
    <%
	else
	response.Write("<tr><td colspan=2 bgcolor=#FFFFFF>输入用户名错误!没有找到此用户</td></tr>")
	end if
	case "step3"
	set rs=conn.execute("select quesion,answer,user from userinfo where user='"&request.form("user")&"'")
	if (not rs.eof) and (not rs.bof) then
	  if rs("answer")=MD5(trim(request.Form("answer"))) then
	%>
	    <tr bgcolor=#FFFFFF> 
      <td height="26" colspan="2" background="backimg/bg1.gif"> <b><font color="#FFFFFF">会员找回密码</font></b> 
        第三步 (<strong>修改密码</strong>)</td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=4>用户名:</td>
      <td><%=rs("user")%> <input name="user" type="hidden" id="action" value="<%=rs("user")%>"> 
        <input name="action" type="hidden" id="action" value="step4"> <input name="answer" type="hidden" id="answer" value="<%=request.Form("answer")%>"> 
      </td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td height=1>问 题:</td>
      <td><%=rs("quesion")%></td>
    </tr>
    <tr bgcolor=#FFFFFF>
      <td height=2>答 案:</td>
      <td><%=request.Form("answer")%></td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td height=9>密 码：</td>
      <td><input name=password1 type=password id=password12 size=30></td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=9>确认密码: </td>
      <td> <input name=password2 type=password id=password1 size=30></td>
    </tr>
    <%
  else
  response.Write("<tr><td colspan=2 bgcolor=#FFFFFF>输入问题答案不正确!请重新输入!</td></tr>")
  end if
else
response.Write("<tr><td colspan=2 bgcolor=#FFFFFF>输入用户名错误!没有找到此用户</td></tr>")
end if
case "step4"
%>
	  <tr bgcolor=#FFFFFF> 
      <td height="26" colspan="2" background="backimg/bg1.gif"> <b><font color="#FFFFFF">会员找回密码</font></b> 
        <strong>(取回密码成功)</strong></td>
    </tr><tr><td height=10 colspan=2 bgcolor=#FFFFFF><%=errmess%><br>
       <li>你已成功的修改了密码.</td>
    </tr>
<%
end select
%>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td height="10" colspan="2">
	  <%if action<>"step4" then%>
	  <input type="submit" name="Submit" value="下一步"> 
	  <%else%>
	   <input type="button" name="Submit" onClick="location='Login.asp'"value="返 回"> 
	 <%end if%>
      </td>
    </tr>
  </table>
</form>
<!--#include file="end.asp" -->
</body>
</html>
