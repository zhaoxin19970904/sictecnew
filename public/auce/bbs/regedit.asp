<!-- #Include File="CONN.ASP"-->
<!--#include file="md5.asp" -->
<!--#include file="ubbcode.asp" -->
<%
if request.Form("form1")<>"" then
call regupdate()
end if
sub regupdate()
'on error resume next
dim Username,Password,Password2,quesion,answer,Email,rs,sql,qmlong,birthday,objitem
Username = Trim(replace(Request.form("user"),"'",""))
Password = Trim(replace(Request.form("Password"),"'",""))
Password2 = Trim(replace(Request.form("Passwordr"),"'",""))
quesion=Trim(request.form("quesion"))
answer=Trim(request.form("answer"))
Email = Trim(Request.form("Email"))
If Username = "" Or Password = ""  or Email = "" or quesion="" or answer="" Then
errmess="\n请检查您填写的内容是否完整"
ElseIf Instr(Email, "@") = 0 Or Right(Email, 1) = "@" Or Left(Email, 1) = "@" Then
errmess=errmess&"\n请检查您的邮件地址是否正确！"
ElseIf Password <> Password2 Then
errmess=errmess&"\n你两次输入的密码不一到致!"
End If

set rs = conn.Execute("Select user From userinfo Where user = '"&Username&"'" )
If Not rs.EOF Then
errmess=errmess&"\n你注册的用户已经存,请更换用户名"
End If
if Len(Password)<6 then
errmess=errmess&"\n请注意!密码位数不能少于6位."
end if
qmlong=split(request.form("Signature"),vbCrlf)
if ubound(qmlong)>8 then 
errmess=errmess&"\n签名不能超过8行！请更正好再提交"
End If
set rs = conn.Execute("Select email From userinfo Where email = '"&email&"'" )
If Not rs.EOF Then
errmess=errmess&"\n此E-mail已有人使用了!请更换E-mail"
End If
if errmess<>"" then
response.write "<script language=JavaScript>" & chr(13) & "alert('"&errmess&"'); history.back() </script>"
response.End()
end if

birthday=request.form("byear")&"-"&request.form("bmonth")&"-"&request.form("bday")
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "userinfo", conn, 1, 3
rs.addnew

rs("user")=Username
rs("password")=MD5(password)
rs("xb")=request.Form("xb")
rs("email")=email
rs("quesion")=quesion
rs("answer")=MD5(answer)
rs("birthday")=birthday
rs("oicq")=request.Form("oicq")
rs("myface")=request.Form("myface")
rs("conn")=dvHTMLEncode(request.Form("conn"))
rs("ubbqm")=userubbqm(request.form("Signature"))
rs("Signature")=request.form("Signature")
rs("userseting")=request.Form("setuser1")&"|"&request.Form("mess")
rs("time")=now()
rs.update
set rs=nothing
response.cookies("renwen")("user")=Username
response.cookies("renwen")("passedok")="ofdkjduy"

conn.execute("update admin_copyc set reguser=reguser+1 where id=1")
conn.execute("update userinfo set lastlogin='"&now()&"',logout=logout+1,lastlogIP='"&request.ServerVariables("REMOTE_ADDR")&"' where user='"&Username&"'")
response.Cookies("rewin")("errmess")="<li>您已经成为数字论坛的注册用户，感谢您的支持！<li>单击这里 <a href=index.asp>进入论坛</a>"
Response.Redirect "error.asp?founderr=mess"
Response.end
end sub
%>
<!--#include file="mymem.asp" -->
<%if request("action")<>"apply" then%>
<table width="898" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
  <tr>
    <td height="27" align="center" background="backimg/bg1.gif" bgcolor="#FFFFFF"><font color="#FFFFFF"><strong>论坛规则说明 
      </strong></font> </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> <form name="form2" method="post" action="regedit.asp?action=apply">

        <p>&nbsp;</p>
        <p>(2000年12月28日第九届全国人民代表大会常务委员会第十九次会议通过)</p>
        <p> <br>
          　　我国的互联网，在国家大力倡导和积极推动下，在经济建设和各项事业中得到日益广泛的应用，使人们的生产、工作、学习和生活方式已经开始并将继续发生深刻的变化，对于加快我国国民经济、科学技术的发展和社会服务信息化进程具有重要作用。同时，如何保障互联网的运行安全和信息安全问题已经引起全社会的普遍关注。为了兴利除弊，促进我国互联网的健康发展，维护国家安全和社会公共利益，保护个人、法人和其他组织的合法权益，特作如下决定：</p>
        <p>　　一、为了保障互联网的运行安全，对有下列行为之一，构成犯罪的，依照刑法有关规定追究刑事责任：</p>
        <p>　　(一)侵入国家事务、国防建设、尖端科学技术领域的计算机信息系统；</p>
        <p>　　(二)故意制作、传播计算机病毒等破坏性程序，攻击计算机系统及通信网络，致使计算机系统及通信网络遭受损害；</p>
        <p>　　(三)违反国家规定，擅自中断计算机网络或者通信服务，造成计算机网络或者通信系统不能正常运行。</p>
        <p>　　二、为了维护国家安全和社会稳定，对有下列行为之一，构成犯罪的，依照刑法有关规定追究刑事责任：</p>
        <p>　　(一)利用互联网造谣、诽谤或者发表、传播其他有害信息，煽动颠覆国家政权、推翻社会主义制度，或者煽动分裂国家、破坏国家统一；</p>
        <p>　　(二)通过互联网窃取、泄露国家秘密、情报或者军事秘密；</p>
        <p>　　(三)利用互联网煽动民族仇恨、民族歧视，破坏民族团结；</p>
        <p>　　(四)利用互联网组织邪教组织、联络邪教组织成员，破坏国家法律、行政法规实施。</p>
        <p>　　三、为了维护社会主义市场经济秩序和社会管理秩序，对有下列行为之一，构成犯罪的，依照刑法有关规定追究刑事责任：(一)利用互联网销售伪劣产品或者对商品、服务作虚假宣传；</p>
        <p>　　(二)利用互联网损害他人商业信誉和商品声誉；</p>
        <p>　　(三)利用互联网侵犯他人知识产权；</p>
        <p>　　(四)利用互联网编造并传播影响证券、期货交易或者其他扰乱金融秩序的虚假信息；</p>
        <p>　　(五)在互联网上建立淫秽网站、网页，提供淫秽站点链接服务，或者传播淫秽书刊、影片、音像、图片。</p>
        <p>　　四、为了保护个人、法人和其他组织的人身、财产等合法权利，对有下列行为之一，构成犯罪的，依照刑法有关规定追究刑事责任：</p>
        <p>　　(一)利用互联网侮辱他人或者捏造事实诽谤他人；</p>
        <p>　　(二)非法截获、篡改、删除他人电子邮件或者其他数据资料，侵犯公民通信自由和通信秘密；</p>
        <p>　　(三)利用互联网进行盗窃、诈骗、敲诈勒索。</p>
        <p>　　五、利用互联网实施本决定第一条、第二条、第三条、第四条所列行为以外的其他行为，构成犯罪的，依照刑法有关规定追究刑事责任。</p>
        <p>　　六、利用互联网实施违法行为，违反社会治安管理，尚不构成犯罪的，由公安机关依照《治安管理处罚条例》予以处罚；违反其他法律、行政法规，尚不构成犯罪的，由有关行政管理部门依法给予行政处罚；对直接负责的主管人员和其他直接责任人员，依法给予行政处分或者纪律处分。</p>
        <p>　　利用互联网侵犯他人合法权益，构成民事侵权的，依法承担民事责任。</p>
        <p>　　七、各级人民政府及有关部门要采取积极措施，在促进互联网的应用和网络技术的普及过程中，重视和支持对网络安全技术的研究和开发，增强网络的安全防护能力。有关主管部门要加强对互联网的运行安全和信息安全的宣传教育，依法实施有效的监督管理，防范和制止利用互联网进行的各种违法活动，为互联网的健康发展创造良好的社会环境。从事互联网业务的单位要依法开展活动，发现互联网上出现违法犯罪行为和有害信息时，要采取措施，停止传输有害信息，并及时向有关机关报告。任何单位和个人在利用互联网时，都要遵纪守法，抵制各种违法犯罪行为和有害信息。人民法院、人民检察院、公安机关、国家安全机关要各司其职，密切配合，依法严厉打击利用互联网实施的各种犯罪活动。要动员全社会的力量，依靠全社会的共同努力，保障互联网的运行安全与信息安全，促进社会主义精神文明和物质文明建设。</p>
        <p>以上内容，请仔细阅读。</p>
        <div align="center"> 
          <p class="x"> 
            <SCRIPT LANGUAGE="JAVASCRIPT">
<!--
var time_start = new Date();
var clock_start = time_start.getSeconds();
function show_secs () 
{ 
var ntime= new Date();
var secs=ntime.getSeconds();
lastsec=((clock_start+5)-secs);
document.form2.Submit.disabled=true;
document.form2.Submit.value = "阅读:"+lastsec ;
window.setTimeout('show_secs()',1000); 
if (lastsec<0 ||lastsec>5){
document.form2.Submit.value = "我同意";
document.form2.Submit.disabled=false;
}
}
show_secs ();
// -->
</SCRIPT>
            <input type="submit" name="Submit" value="我同意">
            <input type="button" name="Submit3" onclick=history.go(-1) value="不同意">
          </p>
        </div>
      </form></td>
  </tr>

</table>
<p></p>
<%else%>
<FORM name=form1 action=regedit.asp?action=apply method=post>
  <table width="770" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
    <tr class="tableborder1"> 
      <td height="27" colspan="2" class=tablebody1 background="backimg/bg1.gif"> <font color="#FFFFFF" size="2"><strong>新用户注册 
        </strong> </font></td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF" class=tablebody1><font color="#FF0000">*</font>用户名：<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input maxlength=16 size=30 name=user> 
        &nbsp;&nbsp; 注册用户名不能超过16个字符（8个汉字） </td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1><font color="#FF0000">*</font>性别：<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input type=radio CHECKED value=男孩 name=xb> 
        <img  src=pic/Male.gif align=absMiddle>男孩 &nbsp;&nbsp;&nbsp;&nbsp; <input type=radio value=女孩 name=xb> 
        <img  src=pic/Female.gif align=absMiddle>女孩&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        &nbsp;请选择您的性别</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF" class=tablebody1><font color="#FF0000">*</font>密码(至少6位)：<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF" class=tablebody1> <input type=password maxlength=16 size=30 name=password> 
        &nbsp;&nbsp;&nbsp; 请输入密码。 请不要使用任何类似 '*'、' ' 或 HTML 字符 </td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF" class=tablebody1><font color="#FF0000">*</font>密码(至少6位)：<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF" class=tablebody1> <input type=password maxlength=16 size=30 name=passwordr> 
        &nbsp;&nbsp;&nbsp; 请再输一遍确认 &nbsp; &nbsp; </td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1><font color="#FF0000">*</font>Email地址：<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input maxlength=50 size=30 name=email>
        请输入有效的邮件地址这将使您能用 到论坛中的所有功能</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>密码问题：</td>
      <td class=tablebody1> <input type=text size=30 name=quesion>
        忘记密码的提示问题 </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>问题答案：<br> </td>
      <td class=tablebody1> <input type=text size=30 name=answer> &nbsp; &nbsp; 
        忘记密码的提示问题答案，用于取回论坛密码 </td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" valign=top bgcolor="#FFFFFF" >自定义头像：<br> </td>
      <td width="595" bgcolor="#FFFFFF"  > <select name=myface 
            onChange="document.images['idface'].src=options[selectedIndex].value;" 
            size=1>
          <option value="face/1.gif" selected>头像1</option>
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
        </select>
        自定头像&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; 请选择你的头像，<font color="#FF0000">重登录修改资料</font>便可自定义头像 
        <div id="Layer1" style="position:absolute; width:23px; height:31px; z-index:1"><img src="face/1.gif" id=idface></div></td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1>生日<br> </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <select style="FONT-SIZE: 9pt" 
             name=byear>
          <script language=JavaScript><!--               
          for(i=1900;i<2005;i++) document.write('<option>'+i)               
            //-->
         </script>
        </select>
        年 
        <select name="bmonth">
          <option value="01" selected>01</option>
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
          <option value="01" selected>01</option>
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
        日&nbsp;&nbsp;&nbsp;&nbsp; 如不想填写，你将是全世界最老的人</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1> OICQ</td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input maxlength=50 size=30 name=oicq> 
        &nbsp;&nbsp;&nbsp; 请填写好你的QQ号，如果没有可以不填</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1>自定义联系方式<br>
        ( 自行输入你的可联系方式.以便以后我们寄出奖品)</td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <textarea name=conn rows=4 wrap=PHYSICAL cols=40 onKeyDown="textCounter(this.form.conn,this.form.remLen2,300);" onKeyUp="textCounter(this.form.conn,this.form.remLen2,300);"></textarea> 
        &nbsp; 联系方式尚能输入 
        <input name=remLen2 type=text style="font-size:9pt;color:red;border-top-width:0;border-left-width:0;border-right-width:0;border-bottom-width:0" value="250" size=3 maxlength=3 readonly>
        个字符</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=x> <p>签名：(支持文字UBB代码.不支持HTML)<br>
        </p></td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <textarea name=Signature rows=5 wrap=PHYSICAL  cols=40 onKeyDown="textCounter(this.form.Signature,this.form.remLen,250);" onKeyUp="textCounter(this.form.Signature,this.form.remLen,300);" ></textarea> 
        <SCRIPT LANGUAGE="JavaScript">
<!--//
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
</SCRIPT>
        签名尚能输入 
        <input name=remLen type=text style="font-size:9pt;color:red;border-top-width:0;border-left-width:0;border-right-width:0;border-bottom-width:0" value="300" size=3 maxlength=3 readonly>
        个字符</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1>是否开放您的基本资料：<br> </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input type=radio name=setuser1 value=1 checked>
        开放 
        <input type=radio name=setuser1 value=0>
        不开放&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 开放后别人可以看到您的性别、Email、QQ等信息</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bordercolor="#CCCCCC" bgcolor="#FFFFFF"  class=tablebody1>是否接受短消息</td>
      <td width="595" bordercolor="#CCCCCC" bgcolor="#FFFFFF"  class=tablebody1> 
        <input name="mess" type="radio" value="0" checked>
        接受 
        <input type="radio" name="mess" value="1">
        拒收</td>
    </tr>
    <tr class="tableborder1"> 
      <td colspan="2" bgcolor="#FFFFFF"  class=tablebody1> <div align="center"> 
          <font size="2"> 
          <input type=submit value="注 册" name=form1>
           
          <input type=reset value="清 除" name=Submit2>
          </font></div></td>
    </tr>
  </table>
</form>
<%end if%>
<!--#include file="end.asp" -->

</body>
</html>
