<!--#include file="conn.asp" -->
<%
if request("form1") <> "" then
dim email,rs,sql,user,email,pass
email=request.form("email")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from user where email='"&email&"'" 
rs.open sql,conn,1,1
 if not rs.eof then
  do while not rs.eof 
user=rs("user")
email=rs("email")
pass=rs("password")
Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.From ="mtv@mtvok.com"
objCDOMail.To =email
objCDOMail.Subject ="一封来自数字论坛论坛的密码查询信"
objCDOMail.BodyFormat = 0 
objCDOMail.MailFormat = 0 
objCDOMail.Body ="<p>"&user&":你好</p><p> 感谢来到数字论坛论坛注册.我们将忠心有你服务.欢迎你提出宝贵的意见.</p><p>你的注册信息为：</p> <p>用户名:"&user&"</p><p>注册E-mail:"&email&"</p><p>密码:"&pass&"</p><p>再次感谢你的支持.欢迎经常光临本站.请记住本站网站:<a href=""http://www.mtvok.com"" target=""_blank"">http://www.mtvok.com</a>本站域名:<a href=""http://www.mtvok.com"" target=""_blank"">http://www.mtvok.com</a></p>"
objCDOMail.Send
Set objCDOMail = Nothing
Response.Write "谢谢你使用本系系.你的注册信息已经成功发送到你的信箱.请注意查收<br><p onClick='history.go(-2)'><a href=""#"">返回</a>&nbsp;</p>"
   rs.movenext
   loop
  else 
response.write"对不起! 我们没有找到你所输入的注册E-mail.看看是不是E-mail填写有误.或是你还没有注册请去 <a href=""regedit.asp"" >注册</a><br><p onClick='history.go(-1)'><a href=""#"">返回</a>&nbsp;</p>"
  end if
Response.End
end if
if request("form2")<>"" then
user=request.form("user")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from user where user='"&user&"'" 
rs.open sql,conn,1,1
 if not rs.eof then
  do while not rs.eof 
user=rs("user")
email=rs("email")
pass=rs("password")
Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.From ="mtv@mtvok.com"
objCDOMail.To =email
objCDOMail.Subject ="一封来自数字论坛论坛的密码查询信"
objCDOMail.BodyFormat = 0 
objCDOMail.MailFormat = 0 
objCDOMail.Body ="<p>"&user&":你好</p><p> 感谢来到数字论坛论坛注册.我们将忠心有你服务.欢迎你提出宝贵的意见.</p><p>你的注册信息为：</p> <p>用户名:"&user&"</p><p>注册E-mail:"&email&"</p><p>密码:"&pass&"</p><p>再次感谢你的支持.欢迎经常光临本站.请记住本站网站:<a href=""http://www.mtvok.com"" target=""_blank"">http://www.mtvok.com</a>本站域名:<a href=""http://www.mtvok.com"" target=""_blank"">http://www.mtvok.com</a></p>"
objCDOMail.Send
Set objCDOMail = Nothing
Response.Write "谢谢你使用本系系.你的注册信息已经成功发送到你的信箱.请注意查收<br><p onClick='history.go(-1)'><a href=""#"">返回</a>&nbsp;</p>"
   rs.movenext
   loop
  else 
response.write"对不起! 我们没有找到你所输入的注册名.看看是不是注册名填写有误.如果你还没有注册请去 <a href=""regedit.asp"" >注册</a><br><p onClick='history.go(-1)'><a href=""#"">返回</a>&nbsp;</p>"
  end if
Response.End
end if
if request.form("user2")<>"" then
user=request.form("user2")
quesion=request.form("quesion")
answer=request.form("answer")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from user where user='"&user&"'" 
rs.open sql,conn,1,1
if rs("quesion")=quesion and rs("answer")=answer then
response.write"你的注册名:"&rs("user")&"<br>注册密码:"&rs("password")&"<br>请住你的密码!"
response.end
else 
response.write"对不起! 我们没有找到你所输入的注册名.或者是问题和答案不正确!.如果你还没有注册请去 <a href=""regedit.asp"" >注册</a><br><p onClick='history.go(-1)'><a href=""#"">返回</a>&nbsp;</p>"
end if 
end if
%>

  <!--#include file="mymem.asp" -->
<div align="center"><font size="5"><strong>找回密码方法一 </strong></font></div>

  <p align="center"><font size="5"><b><font size="5"><b><font size="4">找回密码方法二</font></b></font></b></font></p>


<form name="form1" action="seamail.asp" method="post" >
  <table width="63%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#92b9fb">
    <tr>
      <td>
        <table width="100%" align="center" cellpadding="4" cellspacing="1">
          <tr bgcolor=""> 
            <td height="26" colspan="3" background="backimg/bg1.gif"> 
              <div align="center"> 
                <p align="center"><font size="5"><b><font size="4">会员找回密码</font></b><font size="4"><font size="2">(</font></font><font size="2">凭E-mail.密码会发到你的邮箱)</font></font></p>
              </div></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="14%" height="31">E-mail </td>
            <td width="36%"> 
              <input type="text" name="email">
            </td>
            <td width="50%"> 
              <div align="center">
                <input type="submit" name="FORM1" value="提交">
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>  <p align="center"><font size="5"><b><font size="5"><b><font size="4">找回密码方法三</font></b></font></b></font></p>

<form name="form2" action="seamail.asp" method="post" >
  <table width="63%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#92b9fb">
    <tr> 
      <td>
        <table width="100%" align="center" cellpadding="4" cellspacing="1">
          <tr bgcolor=""> 
            <td height="26" colspan="3" background="backimg/bg1.gif"> 
              <div align="center"> 
                <p align="center"><font size="5"><b><font size="5"><b><font size="4">会员找回密码</font></b></font></b><font size="5"><font size="4"><font size="2">(凭用户名</font></font></font><font size="2">.密码会发到你的邮箱)</font></font></p>
              </div></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="17%" height="41">用户名: </td>
            <td width="34%"> 
              <input name="user" type="text" id="user"></td>
            <td width="49%"> 
              <div align="center">
                <input type="submit" name="FORM2" value="提交">
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>


<!--#include file="end.asp" -->
</body>
</html>




















