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
objCDOMail.Subject ="һ������������̳��̳�������ѯ��"
objCDOMail.BodyFormat = 0 
objCDOMail.MailFormat = 0 
objCDOMail.Body ="<p>"&user&":���</p><p> ��л����������̳��̳ע��.���ǽ������������.��ӭ�������������.</p><p>���ע����ϢΪ��</p> <p>�û���:"&user&"</p><p>ע��E-mail:"&email&"</p><p>����:"&pass&"</p><p>�ٴθ�л���֧��.��ӭ�������ٱ�վ.���ס��վ��վ:<a href=""http://www.mtvok.com"" target=""_blank"">http://www.mtvok.com</a>��վ����:<a href=""http://www.mtvok.com"" target=""_blank"">http://www.mtvok.com</a></p>"
objCDOMail.Send
Set objCDOMail = Nothing
Response.Write "лл��ʹ�ñ�ϵϵ.���ע����Ϣ�Ѿ��ɹ����͵��������.��ע�����<br><p onClick='history.go(-2)'><a href=""#"">����</a>&nbsp;</p>"
   rs.movenext
   loop
  else 
response.write"�Բ���! ����û���ҵ����������ע��E-mail.�����ǲ���E-mail��д����.�����㻹û��ע����ȥ <a href=""regedit.asp"" >ע��</a><br><p onClick='history.go(-1)'><a href=""#"">����</a>&nbsp;</p>"
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
objCDOMail.Subject ="һ������������̳��̳�������ѯ��"
objCDOMail.BodyFormat = 0 
objCDOMail.MailFormat = 0 
objCDOMail.Body ="<p>"&user&":���</p><p> ��л����������̳��̳ע��.���ǽ������������.��ӭ�������������.</p><p>���ע����ϢΪ��</p> <p>�û���:"&user&"</p><p>ע��E-mail:"&email&"</p><p>����:"&pass&"</p><p>�ٴθ�л���֧��.��ӭ�������ٱ�վ.���ס��վ��վ:<a href=""http://www.mtvok.com"" target=""_blank"">http://www.mtvok.com</a>��վ����:<a href=""http://www.mtvok.com"" target=""_blank"">http://www.mtvok.com</a></p>"
objCDOMail.Send
Set objCDOMail = Nothing
Response.Write "лл��ʹ�ñ�ϵϵ.���ע����Ϣ�Ѿ��ɹ����͵��������.��ע�����<br><p onClick='history.go(-1)'><a href=""#"">����</a>&nbsp;</p>"
   rs.movenext
   loop
  else 
response.write"�Բ���! ����û���ҵ����������ע����.�����ǲ���ע������д����.����㻹û��ע����ȥ <a href=""regedit.asp"" >ע��</a><br><p onClick='history.go(-1)'><a href=""#"">����</a>&nbsp;</p>"
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
response.write"���ע����:"&rs("user")&"<br>ע������:"&rs("password")&"<br>��ס�������!"
response.end
else 
response.write"�Բ���! ����û���ҵ����������ע����.����������ʹ𰸲���ȷ!.����㻹û��ע����ȥ <a href=""regedit.asp"" >ע��</a><br><p onClick='history.go(-1)'><a href=""#"">����</a>&nbsp;</p>"
end if 
end if
%>

  <!--#include file="mymem.asp" -->
<div align="center"><font size="5"><strong>�һ����뷽��һ </strong></font></div>

  <p align="center"><font size="5"><b><font size="5"><b><font size="4">�һ����뷽����</font></b></font></b></font></p>


<form name="form1" action="seamail.asp" method="post" >
  <table width="63%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#92b9fb">
    <tr>
      <td>
        <table width="100%" align="center" cellpadding="4" cellspacing="1">
          <tr bgcolor=""> 
            <td height="26" colspan="3" background="backimg/bg1.gif"> 
              <div align="center"> 
                <p align="center"><font size="5"><b><font size="4">��Ա�һ�����</font></b><font size="4"><font size="2">(</font></font><font size="2">ƾE-mail.����ᷢ���������)</font></font></p>
              </div></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="14%" height="31">E-mail </td>
            <td width="36%"> 
              <input type="text" name="email">
            </td>
            <td width="50%"> 
              <div align="center">
                <input type="submit" name="FORM1" value="�ύ">
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>  <p align="center"><font size="5"><b><font size="5"><b><font size="4">�һ����뷽����</font></b></font></b></font></p>

<form name="form2" action="seamail.asp" method="post" >
  <table width="63%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#92b9fb">
    <tr> 
      <td>
        <table width="100%" align="center" cellpadding="4" cellspacing="1">
          <tr bgcolor=""> 
            <td height="26" colspan="3" background="backimg/bg1.gif"> 
              <div align="center"> 
                <p align="center"><font size="5"><b><font size="5"><b><font size="4">��Ա�һ�����</font></b></font></b><font size="5"><font size="4"><font size="2">(ƾ�û���</font></font></font><font size="2">.����ᷢ���������)</font></font></p>
              </div></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="17%" height="41">�û���: </td>
            <td width="34%"> 
              <input name="user" type="text" id="user"></td>
            <td width="49%"> 
              <div align="center">
                <input type="submit" name="FORM2" value="�ύ">
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>


<!--#include file="end.asp" -->
</body>
</html>




















