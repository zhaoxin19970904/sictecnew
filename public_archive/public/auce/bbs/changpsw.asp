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
errmess=errmess&"�û������ڣ���ע���û����Ƿ�������ȫ���ַ��������Ŀո�"
exit sub
end if
If Rs("password")<> MD5(password)  Then
errmess=errmess&"������ľ����벻��ȷ"
exit sub
End If
If password2 <>password3  Then
errmess=errmess&"��������������벻һ�£�"
exit sub
End If
if Len(Password2)<6 then
errmess=errmess&"����λ����������6λ."
exit sub
end if
conn.execute("update userinfo set [password]='"&MD5(password2)&"' where user='"&user&"'")
	founderr=true
    errmess=errmess&"<li><b>�����޸ĳɹ�! </b> ���ס������!"
	errmess=errmess&"<li>�뷵����ҳ�ص�¼.�����Զ�������ҳ.<li><li><a href=index.asp>������ҳ</a>"
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
      <td height="27" colspan="2" background="backimg/bg1.gif" class=link><b><font color="#FFFFFF">�û����ĵ�¼����</font></b></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td colspan="2" class=link><%if errmess<>"" then response.write(errmess) end if%></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td width=223 class=link>�û����ƣ�</td>
      <td> <input name=user type=text class=link id="user2" size=20 maxlength=16> 
        <font color="#FF3333">*</font> ��ĸ������</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td width=223 class=link>�� �� �룺</td>
      <td width=539> <input type=password name=password size=20 maxlength=16 class=link> 
        <font color="#FF3333">*</font> ��ĸ������</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td width=223 class=link>�� �� �룺</td>
      <td width=539> <input type=password name=password2 size=20 maxlength=16 class=link> 
        <font color="#FF3333">*</font> ��ĸ������</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td width=223 class=link>ȷ�����룺</td>
      <td width=539> <input type=password name=password3 size=20 maxlength=16 class=link> 
        <font color="#FF3333">*</font> ��ĸ������</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td colspan="2" class=link> <input type="submit" value="ȷ���޸�" name="b1">
        �� 
        <input type="reset" name="Reset" value="������д"> </td>
    </tr>
	</form>
  </table>


<!--#include file="end.asp" -->
</body>
</html>