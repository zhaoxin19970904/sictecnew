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
errmess="\n��������д�������Ƿ�����"
ElseIf Instr(Email, "@") = 0 Or Right(Email, 1) = "@" Or Left(Email, 1) = "@" Then
errmess=errmess&"\n���������ʼ���ַ�Ƿ���ȷ��"
ElseIf Password <> Password2 Then
errmess=errmess&"\n��������������벻һ����!"
End If

set rs = conn.Execute("Select user From userinfo Where user = '"&Username&"'" )
If Not rs.EOF Then
errmess=errmess&"\n��ע����û��Ѿ���,������û���"
End If
if Len(Password)<6 then
errmess=errmess&"\n��ע��!����λ����������6λ."
end if
qmlong=split(request.form("Signature"),vbCrlf)
if ubound(qmlong)>8 then 
errmess=errmess&"\nǩ�����ܳ���8�У�����������ύ"
End If
set rs = conn.Execute("Select email From userinfo Where email = '"&email&"'" )
If Not rs.EOF Then
errmess=errmess&"\n��E-mail������ʹ����!�����E-mail"
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
response.Cookies("rewin")("errmess")="<li>���Ѿ���Ϊ������̳��ע���û�����л����֧�֣�<li>�������� <a href=index.asp>������̳</a>"
Response.Redirect "error.asp?founderr=mess"
Response.end
end sub
%>
<!--#include file="mymem.asp" -->
<%if request("action")<>"apply" then%>
<table width="898" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
  <tr>
    <td height="27" align="center" background="backimg/bg1.gif" bgcolor="#FFFFFF"><font color="#FFFFFF"><strong>��̳����˵�� 
      </strong></font> </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> <form name="form2" method="post" action="regedit.asp?action=apply">

        <p>&nbsp;</p>
        <p>(2000��12��28�յھŽ�ȫ����������᳣��ίԱ���ʮ�Ŵλ���ͨ��)</p>
        <p> <br>
          �����ҹ��Ļ��������ڹ��Ҵ��������ͻ����ƶ��£��ھ��ý���͸�����ҵ�еõ�����㷺��Ӧ�ã�ʹ���ǵ�������������ѧϰ�����ʽ�Ѿ���ʼ��������������̵ı仯�����ڼӿ��ҹ����񾭼á���ѧ�����ķ�չ����������Ϣ�����̾�����Ҫ���á�ͬʱ����α��ϻ����������а�ȫ����Ϣ��ȫ�����Ѿ�����ȫ�����ձ��ע��Ϊ���������ף��ٽ��ҹ��������Ľ�����չ��ά�����Ұ�ȫ����ṫ�����棬�������ˡ����˺�������֯�ĺϷ�Ȩ�棬�������¾�����</p>
        <p>����һ��Ϊ�˱��ϻ����������а�ȫ������������Ϊ֮һ�����ɷ���ģ������̷��йع涨׷���������Σ�</p>
        <p>����(һ)����������񡢹������衢��˿�ѧ��������ļ������Ϣϵͳ��</p>
        <p>����(��)��������������������������ƻ��Գ��򣬹��������ϵͳ��ͨ�����磬��ʹ�����ϵͳ��ͨ�����������𺦣�</p>
        <p>����(��)Υ�����ҹ涨�������жϼ�����������ͨ�ŷ�����ɼ�����������ͨ��ϵͳ�����������С�</p>
        <p>��������Ϊ��ά�����Ұ�ȫ������ȶ�������������Ϊ֮һ�����ɷ���ģ������̷��йع涨׷���������Σ�</p>
        <p>����(һ)���û�������ҥ���̰����߷������������к���Ϣ��ɿ���߸�������Ȩ���Ʒ���������ƶȣ�����ɿ�����ѹ��ҡ��ƻ�����ͳһ��</p>
        <p>����(��)ͨ����������ȡ��й¶�������ܡ��鱨���߾������ܣ�</p>
        <p>����(��)���û�����ɿ�������ޡ��������ӣ��ƻ������Ž᣻</p>
        <p>����(��)���û�������֯а����֯������а����֯��Ա���ƻ����ҷ��ɡ���������ʵʩ��</p>
        <p>��������Ϊ��ά����������г�������������������򣬶���������Ϊ֮һ�����ɷ���ģ������̷��йع涨׷���������Σ�(һ)���û���������α�Ӳ�Ʒ���߶���Ʒ�����������������</p>
        <p>����(��)���û�������������ҵ��������Ʒ������</p>
        <p>����(��)���û������ַ�����֪ʶ��Ȩ��</p>
        <p>����(��)���û��������첢����Ӱ��֤ȯ���ڻ����׻����������ҽ�������������Ϣ��</p>
        <p>����(��)�ڻ������Ͻ���������վ����ҳ���ṩ����վ�����ӷ��񣬻��ߴ��������鿯��ӰƬ������ͼƬ��</p>
        <p>�����ġ�Ϊ�˱������ˡ����˺�������֯�������Ʋ��ȺϷ�Ȩ��������������Ϊ֮һ�����ɷ���ģ������̷��йع涨׷���������Σ�</p>
        <p>����(һ)���û������������˻���������ʵ�̰����ˣ�</p>
        <p>����(��)�Ƿ��ػ񡢴۸ġ�ɾ�����˵����ʼ����������������ϣ��ַ�����ͨ�����ɺ�ͨ�����ܣ�</p>
        <p>����(��)���û��������е��ԡ�թƭ����թ������</p>
        <p>�����塢���û�����ʵʩ��������һ�����ڶ�������������������������Ϊ�����������Ϊ�����ɷ���ģ������̷��йع涨׷���������Ρ�</p>
        <p>�����������û�����ʵʩΥ����Ϊ��Υ������ΰ������в����ɷ���ģ��ɹ����������ա��ΰ����������������Դ�����Υ���������ɡ��������棬�в����ɷ���ģ����й�������������������������������ֱ�Ӹ����������Ա������ֱ��������Ա�����������������ֻ��߼��ɴ��֡�</p>
        <p>�������û������ַ����˺Ϸ�Ȩ�棬����������Ȩ�ģ������е��������Ρ�</p>
        <p>�����ߡ����������������йز���Ҫ��ȡ������ʩ���ڴٽ���������Ӧ�ú����缼�����ռ������У����Ӻ�֧�ֶ����簲ȫ�������о��Ϳ�������ǿ����İ�ȫ�����������й����ܲ���Ҫ��ǿ�Ի����������а�ȫ����Ϣ��ȫ����������������ʵʩ��Ч�ļල������������ֹ���û��������еĸ���Υ�����Ϊ�������Ľ�����չ�������õ���ỷ�������»�����ҵ��ĵ�λҪ������չ������ֻ������ϳ���Υ��������Ϊ���к���Ϣʱ��Ҫ��ȡ��ʩ��ֹͣ�����к���Ϣ������ʱ���йػ��ر��档�κε�λ�͸��������û�����ʱ����Ҫ����ط������Ƹ���Υ��������Ϊ���к���Ϣ������Ժ��������Ժ���������ء����Ұ�ȫ����Ҫ��˾��ְ��������ϣ���������������û�����ʵʩ�ĸ��ַ�����Ҫ��Աȫ��������������ȫ���Ĺ�ͬŬ�������ϻ����������а�ȫ����Ϣ��ȫ���ٽ�������徫�������������������衣</p>
        <p>�������ݣ�����ϸ�Ķ���</p>
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
document.form2.Submit.value = "�Ķ�:"+lastsec ;
window.setTimeout('show_secs()',1000); 
if (lastsec<0 ||lastsec>5){
document.form2.Submit.value = "��ͬ��";
document.form2.Submit.disabled=false;
}
}
show_secs ();
// -->
</SCRIPT>
            <input type="submit" name="Submit" value="��ͬ��">
            <input type="button" name="Submit3" onclick=history.go(-1) value="��ͬ��">
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
      <td height="27" colspan="2" class=tablebody1 background="backimg/bg1.gif"> <font color="#FFFFFF" size="2"><strong>���û�ע�� 
        </strong> </font></td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF" class=tablebody1><font color="#FF0000">*</font>�û�����<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input maxlength=16 size=30 name=user> 
        &nbsp;&nbsp; ע���û������ܳ���16���ַ���8�����֣� </td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1><font color="#FF0000">*</font>�Ա�<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input type=radio CHECKED value=�к� name=xb> 
        <img  src=pic/Male.gif align=absMiddle>�к� &nbsp;&nbsp;&nbsp;&nbsp; <input type=radio value=Ů�� name=xb> 
        <img  src=pic/Female.gif align=absMiddle>Ů��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        &nbsp;��ѡ�������Ա�</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF" class=tablebody1><font color="#FF0000">*</font>����(����6λ)��<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF" class=tablebody1> <input type=password maxlength=16 size=30 name=password> 
        &nbsp;&nbsp;&nbsp; ���������롣 �벻Ҫʹ���κ����� '*'��' ' �� HTML �ַ� </td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF" class=tablebody1><font color="#FF0000">*</font>����(����6λ)��<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF" class=tablebody1> <input type=password maxlength=16 size=30 name=passwordr> 
        &nbsp;&nbsp;&nbsp; ������һ��ȷ�� &nbsp; &nbsp; </td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1><font color="#FF0000">*</font>Email��ַ��<br> 
      </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input maxlength=50 size=30 name=email>
        ����������Ч���ʼ���ַ�⽫ʹ������ ����̳�е����й���</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>�������⣺</td>
      <td class=tablebody1> <input type=text size=30 name=quesion>
        �����������ʾ���� </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tableborder1"> 
      <td  class=tablebody1><font color="#FF0000">*</font>����𰸣�<br> </td>
      <td class=tablebody1> <input type=text size=30 name=answer> &nbsp; &nbsp; 
        �����������ʾ����𰸣�����ȡ����̳���� </td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" valign=top bgcolor="#FFFFFF" >�Զ���ͷ��<br> </td>
      <td width="595" bgcolor="#FFFFFF"  > <select name=myface 
            onChange="document.images['idface'].src=options[selectedIndex].value;" 
            size=1>
          <option value="face/1.gif" selected>ͷ��1</option>
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
        </select>
        �Զ�ͷ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; ��ѡ�����ͷ��<font color="#FF0000">�ص�¼�޸�����</font>����Զ���ͷ�� 
        <div id="Layer1" style="position:absolute; width:23px; height:31px; z-index:1"><img src="face/1.gif" id=idface></div></td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1>����<br> </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <select style="FONT-SIZE: 9pt" 
             name=byear>
          <script language=JavaScript><!--               
          for(i=1900;i<2005;i++) document.write('<option>'+i)               
            //-->
         </script>
        </select>
        �� 
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
        �� 
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
        ��&nbsp;&nbsp;&nbsp;&nbsp; �粻����д���㽫��ȫ�������ϵ���</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1> OICQ</td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input maxlength=50 size=30 name=oicq> 
        &nbsp;&nbsp;&nbsp; ����д�����QQ�ţ����û�п��Բ���</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1>�Զ�����ϵ��ʽ<br>
        ( ����������Ŀ���ϵ��ʽ.�Ա��Ժ����Ǽĳ���Ʒ)</td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <textarea name=conn rows=4 wrap=PHYSICAL cols=40 onKeyDown="textCounter(this.form.conn,this.form.remLen2,300);" onKeyUp="textCounter(this.form.conn,this.form.remLen2,300);"></textarea> 
        &nbsp; ��ϵ��ʽ�������� 
        <input name=remLen2 type=text style="font-size:9pt;color:red;border-top-width:0;border-left-width:0;border-right-width:0;border-bottom-width:0" value="250" size=3 maxlength=3 readonly>
        ���ַ�</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=x> <p>ǩ����(֧������UBB����.��֧��HTML)<br>
        </p></td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <textarea name=Signature rows=5 wrap=PHYSICAL  cols=40 onKeyDown="textCounter(this.form.Signature,this.form.remLen,250);" onKeyUp="textCounter(this.form.Signature,this.form.remLen,300);" ></textarea> 
        <SCRIPT LANGUAGE="JavaScript">
<!--//
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
</SCRIPT>
        ǩ���������� 
        <input name=remLen type=text style="font-size:9pt;color:red;border-top-width:0;border-left-width:0;border-right-width:0;border-bottom-width:0" value="300" size=3 maxlength=3 readonly>
        ���ַ�</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bgcolor="#FFFFFF"  class=tablebody1>�Ƿ񿪷����Ļ������ϣ�<br> </td>
      <td width="595" bgcolor="#FFFFFF"  class=tablebody1> <input type=radio name=setuser1 value=1 checked>
        ���� 
        <input type=radio name=setuser1 value=0>
        ������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ���ź���˿��Կ��������Ա�Email��QQ����Ϣ</td>
    </tr>
    <tr class="tableborder1"> 
      <td width="175" bordercolor="#CCCCCC" bgcolor="#FFFFFF"  class=tablebody1>�Ƿ���ܶ���Ϣ</td>
      <td width="595" bordercolor="#CCCCCC" bgcolor="#FFFFFF"  class=tablebody1> 
        <input name="mess" type="radio" value="0" checked>
        ���� 
        <input type="radio" name="mess" value="1">
        ����</td>
    </tr>
    <tr class="tableborder1"> 
      <td colspan="2" bgcolor="#FFFFFF"  class=tablebody1> <div align="center"> 
          <font size="2"> 
          <input type=submit value="ע ��" name=form1>
          �� 
          <input type=reset value="�� ��" name=Submit2>
          </font></div></td>
    </tr>
  </table>
</form>
<%end if%>
<!--#include file="end.asp" -->

</body>
</html>
