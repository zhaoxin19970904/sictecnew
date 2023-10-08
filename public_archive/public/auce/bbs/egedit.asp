<!--#include file="CONN.ASP" -->
<!--#include file="ubbcode.asp" -->
<!--#include file="md5.asp" -->
<%
if not userislogin then
founderr=true
errmess=errmess+"<li>对不起!我们发现你现在还没登录!所以无法完成你的操作请求!请<a href=Login.asp>登录</a>"
end if
if userloginlock=1 then
founderr=true
errmess=errmess+"<li>对不起!我们发现你的ID被封闭!<li>原因:可能你违反有关规定<li>所以无法完成你的操作请求!请<a href=index.asp>返回首页</a>"
end if
if request.Form("form1")<>"" then
call editupdate()
end if

sub editupdate()
'on error resume next
dim quesion,answer,Email,rs,sql,qmlong,birthday,objitem
quesion=Trim(request.form("quesion"))
answer=Trim(request.form("answer"))
Email = Trim(Request.form("Email"))
If Email = "" or quesion="" or answer="" Then
response.write "<script language=JavaScript>" & chr(13) & "alert('请检查您填写的内容是否完整！');history.back()</script>"
Response.End
end if
If Instr(Email, "@") = 0 Or Right(Email, 1) = "@" Or Left(Email, 1) = "@" Then
response.write "<script language=JavaScript>" & chr(13) & "alert('请检查您的邮件地址是否正确！');history.back();</script>"
Response.End
end if
qmlong=split(request.form("Signature"),vbCrlf)
if ubound(qmlong)>8 then 
response.write "<script language=JavaScript>" & chr(13) & "alert('签名不能超过8行！请更正好再提交');history.back();</script>"
Response.End
End If

if founderr then
exit sub
end if
birthday=request.form("byear")&"-"&request.form("bmonth")&"-"&request.form("bday")
set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from userinfo where user='"&loginuser&"'"
rs.Open sql,conn,1,3
rs("quesion")=quesion
rs("answer")=MD5(answer)
rs("xb")=request.Form("xb")
rs("width")=request.Form("width")
rs("height")=request.Form("height")
rs("email")=email
rs("myface")=request.Form("myface")
rs("conn")=dvHTMLEncode(request.Form("conn"))
rs("ubbqm")=userubbqm(request.form("Signature"))
rs("Signature")=request.form("Signature")
rs("birthday")=birthday
rs("userseting")=request.Form("setuser1")&"|"&request.Form("mess")
rs.update
response.cookies("renwen")("user")=rs("user")
response.cookies("renwen")("passedok")="ofdkjduy"
response.Cookies("rewin")("errmess")="<li>您注册资料已经成功修改，感谢您的支持！单击这里<li><a href=index.asp>进入论坛</a>"
Response.Redirect "error.asp?founderr=mess"
Response.end
end sub
if founderr then
call founderror(errmess)
end if
%>
<!--#include file="mymem.asp" -->
<!--#include file="mymeumu.asp" -->
<%
dim rs,birthday,byear,bday,bmonth
set rs=conn.execute("select * from userinfo where user='"&loginuser&"'")
birthday=rs("birthday")
byear=year(birthday)
bday=day(birthday)
bmonth=month(birthday)
%>
<FORM name=form1  action=egedit.asp method=post >
  <table width="898" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
    <tr class="tableborder1"> 
      <td height="22" colspan="2" background="backimg/bg1.gif" class=tablebody1> 
        <font color="#FFFFFF"><strong>用户资料修改 </strong></font></td>
    </tr>
    <tr bgcolor="#E0E4FE" class="tableborder1"> 
      <td width="19%" bgcolor="#FFFFFF"  class=tablebody1>性别：<br> </td>
      <td width="81%" bgcolor="#FFFFFF"  class=tablebody1> <input type=radio CHECKED value=男孩 name=xb> 
        <img  src=pic/Male.gif align=absMiddle>男孩 &nbsp;&nbsp;&nbsp;&nbsp; <input type=radio value=女孩 name=xb> 
        <img  src=pic/Female.gif align=absMiddle>女孩&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp; 
        请选择您的性别</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>密码问题：</td>
      <td class=tablebody1> <input name=quesion type=text value="<%=rs("quesion")%>" size=30>
        忘记密码的提示问题 </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>问题答案：<br> </td>
      <td class=tablebody1> <input name=answer type=text size=30> 
        &nbsp; &nbsp; 忘记密码的提示问题答案，用于取回论坛密码 </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>Email地址：<br> </td>
      <td  class=tablebody1> <input name=email value="<%=rs("email")%>" size=30 maxlength=50>
        &nbsp;&nbsp; 请输入有效的邮件地址，这将使您能用 到论坛中的所有功能</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td valign=top class=tablebody1>自定义头像：<br>
        允许文件类型:<br>
        (gif,jpg,bmp,jpeg,png) <br> </td>
      <td  class=tablebody1> <iframe name=ad frameborder=0 width=270 height=40 scrolling=no src=upfilex.asp?lb=face></iframe> 
        <br>
        图像位置： 
        <input type=TEXT name=myface size=25 maxlength=100 value="<%=rs("myface")%>"> 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 可以外挂像(请输入完整Url地址) 
        <div id="Layer1" style="position:absolute; width:100px; height:94px; z-index:1;"><img src="../<%=rs("myface")%>" width="<%=rs("width")%>" height="<%=rs("height")%>" id=idface></div>
        <br>
        宽&nbsp;&nbsp;&nbsp;&nbsp;度： 
        <input type=TEXT name=width size=3 maxlength=3 value=<%=rs("width")%>>
        20---120的整数 <br>
        高&nbsp;&nbsp;&nbsp;&nbsp;度： 
        <input type=TEXT name=height size=3 maxlength=3 value=<%=rs("height")%>>
        20---120的整数 <br> <select name=myface1 
            onChange="document.images['idface'].src=options[selectedIndex].value;document.form1.myface.value=myface1.value" 
            size=1>
          <option value="<%=rs("myface")%>" >自定</option>
          <option value="face/1.gif" >头像1</option>
          <option value="face/24.gif">头像2</option>
          <option value="face/2.gif">头像3</option>
          <option value="face/3.gif">头像4</option>
          <option value="face/4.gif">头像5</option>
          <option value="face/5.gif">头像6</option>
          <option value="face/6.gif">头像7</option>
          <option value="face/7.gif">头像8</option>
          <option value="face/8.gif">头像9</option>
          <option value="face/9.gif">头像10</option>
          <option value="face/10.gif">头像11</option>
          <option value="face/11.gif">头像12</option>
          <option value="face/12.gif">头像13</option>
          <option value="face/13.gif">头像14</option>
          <option value="face/14.gif">头像15</option>
          <option value="face/15.gif">头像16</option>
          <option value="face/16.gif">头像17</option>
          <option value="face/17.gif">头像18</option>
          <option value="face/18.gif">头像19</option>
          <option value="face/19.gif">头像20</option>
          <option value="face/20.gif">头像21</option>
          <option value="face/21.gif">头像22</option>
          <option value="face/22.gif">头像23</option>
          <option value="face/23.gif">头像24</option>
        </select> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;如果图像位置中有连接图片将以自定义的为主 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1>生日<br> </td>
      <td  class=tablebody1> <select style="FONT-SIZE: 9pt" 
            onChange=changeCld() name=byear>
          <option value="<%=byear%>" selected><%=byear%></option>
          <script language=JavaScript><!--               
          for(i=1900;i<2005;i++) document.write('<option>'+i)               
            //-->
         </script>
        </select>
        年 
        <select name="bmonth">
          <option value="<%=bmonth%>" selected><%=bmonth%></option>
          <option value="01">01</option>
          <option value="02">02</option>
          <option value="03">03</option>
          <option value="04">04</option>
          <option value="05">05</option>
          <option value="06">06</option>
          <option value="07">07</option>
          <option value="08">08</option>
          <option value="09">09</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
        </select>
        月 
        <select name="bday">
          <option value="<%=bday%>" selected><%=bday%></option>
          <option value="01">01</option>
          <option value="02">02</option>
          <option value="03">03</option>
          <option value="04">04</option>
          <option value="05">05</option>
          <option value="06">06</option>
          <option value="07">07</option>
          <option value="08">08</option>
          <option value="09">09</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="13">13</option>
          <option value="14">14</option>
          <option value="15">15</option>
          <option value="16">16</option>
          <option value="17">17</option>
          <option value="18">18</option>
          <option value="19">19</option>
          <option value="20">20</option>
          <option value="21">21</option>
          <option value="22">22</option>
          <option value="23">23</option>
          <option value="24">24</option>
          <option value="25">25</option>
          <option value="26">26</option>
          <option value="27">27</option>
          <option value="28">28</option>
          <option value="29">29</option>
          <option value="30">30</option>
          <option value="31">31</option>
        </select>
        日&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 如你不填写，你将是世界上最年老的人</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1> OICQ</td>
      <td  class=tablebody1> <input name=oicq value="<%=rs("oicq")%>" size=30 maxlength=50> 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请输入你的OICQ号，如没有可以不用输入</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1>自定义联系方式<br>
        ( 自行输入你的可联系方式.以便以后我们寄出奖品)</td>
      <td  class=tablebody1> <textarea name=conn rows=4 wrap=PHYSICAL cols=40 onKeyDown="textCounter(this.form.conn,this.form.remLen2,300);" onKeyUp="textCounter(this.form.conn,this.form.remLen2,250);"><%=rs("conn")%></textarea> 
        &nbsp; 联系方式尚能输入 
        <input name=remLen2 type=text style="font-size:9pt;color:red;border-top-width:0;border-left-width:0;border-right-width:0;border-bottom-width:0" value="250" size=3 maxlength=3 readonly>
        个字符 </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=x>签名：(支持文字UBB代码,不支持HTML)<br>
        最多300字节文字将出现在您发表<br>
        的文章的结尾处。体现您的个性<br>
        请不要输入太长的空格.</td>
      <td  class=tablebody1> <textarea name=Signature rows=5 wrap=PHYSICAL  cols=40 onKeyDown="textCounter(this.form.Signature,this.form.remLen,300);" onKeyUp="textCounter(this.form.Signature,this.form.remLen,300);" ><%=rs("Signature")%></textarea> 
        <script language="JavaScript">
//-->
function textCounter(field, countfield, maxlimit) {
// 定义函数，传入3个参数，分别为表单区的名字，表单域元素名，字符限制；
if (field.value.length > maxlimit) 
//如果元素区字符数大于最大字符数，按照最大字符数截断； 
field.value = field.value.substring(0, maxlimit);
else
//在记数区文本框内显示剩余的字符数； 
countfield.value = maxlimit - field.value.length;
}
//-->
</script>
        签名尚能输入 
        <input name=remLen type=text style="font-size:9pt;color:red;border-top-width:0;border-left-width:0;border-right-width:0;border-bottom-width:0" value="300" size=3 maxlength=3 readonly>
        个字符 </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1>是否开放您的基本资料：<br> </td>
      <td  class=tablebody1> <input type=radio name=setuser1 value=1 checked>
        开放 
        <input type=radio name=setuser1 value=0>
        不开放&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 开放后别人可以看到您的性别、Email、QQ等信息</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1>是否接受短消息</td>
      <td  class=tablebody1> <input name="mess" type="radio" value="0" checked>
        接受 
        <input type="radio" name="mess" value="1">
        拒收</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td colspan="2"  class=tablebody1> <div align="center"> 
          <input name=form1 type=submit id="form1" onClick="return chkwh()" value="修 改">
           
          <input name=form1 type=reset id="form1" value="清 除">
        </div></td>
    </tr>
  </table>
</form> 
<!--#include file="end.asp" -->
</body>
</html>
