<!--#include file="CONN.ASP" -->
<!--#include file="ubbcode.asp" -->
<!--#include file="md5.asp" -->
<%
if not userislogin then
founderr=true
errmess=errmess+"<li>�Բ���!���Ƿ��������ڻ�û��¼!�����޷������Ĳ�������!��<a href=Login.asp>��¼</a>"
end if
if userloginlock=1 then
founderr=true
errmess=errmess+"<li>�Բ���!���Ƿ������ID�����!<li>ԭ��:������Υ���йع涨<li>�����޷������Ĳ�������!��<a href=index.asp>������ҳ</a>"
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
response.write "<script language=JavaScript>" & chr(13) & "alert('��������д�������Ƿ�������');history.back()</script>"
Response.End
end if
If Instr(Email, "@") = 0 Or Right(Email, 1) = "@" Or Left(Email, 1) = "@" Then
response.write "<script language=JavaScript>" & chr(13) & "alert('���������ʼ���ַ�Ƿ���ȷ��');history.back();</script>"
Response.End
end if
qmlong=split(request.form("Signature"),vbCrlf)
if ubound(qmlong)>8 then 
response.write "<script language=JavaScript>" & chr(13) & "alert('ǩ�����ܳ���8�У�����������ύ');history.back();</script>"
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
response.Cookies("rewin")("errmess")="<li>��ע�������Ѿ��ɹ��޸ģ���л����֧�֣���������<li><a href=index.asp>������̳</a>"
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
        <font color="#FFFFFF"><strong>�û������޸� </strong></font></td>
    </tr>
    <tr bgcolor="#E0E4FE" class="tableborder1"> 
      <td width="19%" bgcolor="#FFFFFF"  class=tablebody1>�Ա�<br> </td>
      <td width="81%" bgcolor="#FFFFFF"  class=tablebody1> <input type=radio CHECKED value=�к� name=xb> 
        <img  src=pic/Male.gif align=absMiddle>�к� &nbsp;&nbsp;&nbsp;&nbsp; <input type=radio value=Ů�� name=xb> 
        <img  src=pic/Female.gif align=absMiddle>Ů��&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp; 
        ��ѡ�������Ա�</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>�������⣺</td>
      <td class=tablebody1> <input name=quesion type=text value="<%=rs("quesion")%>" size=30>
        �����������ʾ���� </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>����𰸣�<br> </td>
      <td class=tablebody1> <input name=answer type=text size=30> 
        &nbsp; &nbsp; �����������ʾ����𰸣�����ȡ����̳���� </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>Email��ַ��<br> </td>
      <td  class=tablebody1> <input name=email value="<%=rs("email")%>" size=30 maxlength=50>
        ����&nbsp;&nbsp; ��������Ч���ʼ���ַ���⽫ʹ������ ����̳�е����й���</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td valign=top class=tablebody1>�Զ���ͷ��<br>
        �����ļ�����:<br>
        (gif,jpg,bmp,jpeg,png) <br> </td>
      <td  class=tablebody1> <iframe name=ad frameborder=0 width=270 height=40 scrolling=no src=upfilex.asp?lb=face></iframe> 
        <br>
        ͼ��λ�ã� 
        <input type=TEXT name=myface size=25 maxlength=100 value="<%=rs("myface")%>"> 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ���������(����������Url��ַ) 
        <div id="Layer1" style="position:absolute; width:100px; height:94px; z-index:1;"><img src="../<%=rs("myface")%>" width="<%=rs("width")%>" height="<%=rs("height")%>" id=idface></div>
        <br>
        ��&nbsp;&nbsp;&nbsp;&nbsp;�ȣ� 
        <input type=TEXT name=width size=3 maxlength=3 value=<%=rs("width")%>>
        20---120������ <br>
        ��&nbsp;&nbsp;&nbsp;&nbsp;�ȣ� 
        <input type=TEXT name=height size=3 maxlength=3 value=<%=rs("height")%>>
        20---120������ <br> <select name=myface1 
            onChange="document.images['idface'].src=options[selectedIndex].value;document.form1.myface.value=myface1.value" 
            size=1>
          <option value="<%=rs("myface")%>" >�Զ�</option>
          <option value="face/1.gif" >ͷ��1</option>
          <option value="face/24.gif">ͷ��2</option>
          <option value="face/2.gif">ͷ��3</option>
          <option value="face/3.gif">ͷ��4</option>
          <option value="face/4.gif">ͷ��5</option>
          <option value="face/5.gif">ͷ��6</option>
          <option value="face/6.gif">ͷ��7</option>
          <option value="face/7.gif">ͷ��8</option>
          <option value="face/8.gif">ͷ��9</option>
          <option value="face/9.gif">ͷ��10</option>
          <option value="face/10.gif">ͷ��11</option>
          <option value="face/11.gif">ͷ��12</option>
          <option value="face/12.gif">ͷ��13</option>
          <option value="face/13.gif">ͷ��14</option>
          <option value="face/14.gif">ͷ��15</option>
          <option value="face/15.gif">ͷ��16</option>
          <option value="face/16.gif">ͷ��17</option>
          <option value="face/17.gif">ͷ��18</option>
          <option value="face/18.gif">ͷ��19</option>
          <option value="face/19.gif">ͷ��20</option>
          <option value="face/20.gif">ͷ��21</option>
          <option value="face/21.gif">ͷ��22</option>
          <option value="face/22.gif">ͷ��23</option>
          <option value="face/23.gif">ͷ��24</option>
        </select> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;���ͼ��λ����������ͼƬ�����Զ����Ϊ�� 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1>����<br> </td>
      <td  class=tablebody1> <select style="FONT-SIZE: 9pt" 
            onChange=changeCld() name=byear>
          <option value="<%=byear%>" selected><%=byear%></option>
          <script language=JavaScript><!--               
          for(i=1900;i<2005;i++) document.write('<option>'+i)               
            //-->
         </script>
        </select>
        �� 
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
        �� 
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
        ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ���㲻��д���㽫�������������ϵ���</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1> OICQ</td>
      <td  class=tablebody1> <input name=oicq value="<%=rs("oicq")%>" size=30 maxlength=50> 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ���������OICQ�ţ���û�п��Բ�������</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1>�Զ�����ϵ��ʽ<br>
        ( ����������Ŀ���ϵ��ʽ.�Ա��Ժ����Ǽĳ���Ʒ)</td>
      <td  class=tablebody1> <textarea name=conn rows=4 wrap=PHYSICAL cols=40 onKeyDown="textCounter(this.form.conn,this.form.remLen2,300);" onKeyUp="textCounter(this.form.conn,this.form.remLen2,250);"><%=rs("conn")%></textarea> 
        &nbsp; ��ϵ��ʽ�������� 
        <input name=remLen2 type=text style="font-size:9pt;color:red;border-top-width:0;border-left-width:0;border-right-width:0;border-bottom-width:0" value="250" size=3 maxlength=3 readonly>
        ���ַ� </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=x>ǩ����(֧������UBB����,��֧��HTML)<br>
        ���300�ֽ����ֽ�������������<br>
        �����µĽ�β�����������ĸ���<br>
        �벻Ҫ����̫���Ŀո�.</td>
      <td  class=tablebody1> <textarea name=Signature rows=5 wrap=PHYSICAL  cols=40 onKeyDown="textCounter(this.form.Signature,this.form.remLen,300);" onKeyUp="textCounter(this.form.Signature,this.form.remLen,300);" ><%=rs("Signature")%></textarea> 
        <script language="JavaScript">
//-->
function textCounter(field, countfield, maxlimit) {
// ���庯��������3���������ֱ�Ϊ���������֣�����Ԫ�������ַ����ƣ�
if (field.value.length > maxlimit) 
//���Ԫ�����ַ�����������ַ�������������ַ����ضϣ� 
field.value = field.value.substring(0, maxlimit);
else
//�ڼ������ı�������ʾʣ����ַ����� 
countfield.value = maxlimit - field.value.length;
}
//-->
</script>
        ǩ���������� 
        <input name=remLen type=text style="font-size:9pt;color:red;border-top-width:0;border-left-width:0;border-right-width:0;border-bottom-width:0" value="300" size=3 maxlength=3 readonly>
        ���ַ� </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1>�Ƿ񿪷����Ļ������ϣ�<br> </td>
      <td  class=tablebody1> <input type=radio name=setuser1 value=1 checked>
        ���� 
        <input type=radio name=setuser1 value=0>
        ������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ���ź���˿��Կ��������Ա�Email��QQ����Ϣ</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1>�Ƿ���ܶ���Ϣ</td>
      <td  class=tablebody1> <input name="mess" type="radio" value="0" checked>
        ���� 
        <input type="radio" name="mess" value="1">
        ����</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td colspan="2"  class=tablebody1> <div align="center"> 
          <input name=form1 type=submit id="form1" onClick="return chkwh()" value="�� ��">
          �� 
          <input name=form1 type=reset id="form1" value="�� ��">
        </div></td>
    </tr>
  </table>
</form> 
<!--#include file="end.asp" -->
</body>
</html>
