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
errmess=errmess&"<li>��������������벻һ��,����������"
exit sub
end if
if Len(password1)<6 then
errmess=errmess&"<li>���벻������6λ�ַ�,����������"
exit sub
end if
if request.Form("user")="" then
errmess=errmess&"<li>�û�����Ϊ��,����������"
exit sub
end if
set rs=conn.execute("select quesion,answer,user from userinfo where user='"&user&"'")
if rs.eof or rs.bof then
errmss=errmess&"<li>���ִ���,�����û���������.����������."
exit sub
end if
if rs("answer")=MD5(trim(request.Form("answer"))) then
conn.execute("update userinfo set [password]='"&MD5(password2)&"' where user='"&user&"'")
errmess=errmess&"<li><b>�����Ѿ��޸ĳɹ�!</b> �����µ�¼,����ס���������"
else
errmess=errmess&"<li>���ִ���,����������𰸲���ȷ.����������."
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
      <td height="26" colspan="2" background="backimg/bg1.gif"> <b><font color="#FFFFFF">��Ա�һ�����</font></b> 
        ��һ�� (<strong>�����û���</strong>)</td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=9>�û���: </td>
      <td><input name="user" type="text" id="user"> <input name="action" type="hidden" id="action" value="step2"> 
      </td>
    </tr>
    <%
	case "step2"
	set rs=conn.execute("select quesion,answer,user from userinfo where user='"&request.form("user")&"'")
	if (not rs.eof) and (not rs.bof) then
	%>
	    <tr bgcolor=#FFFFFF> 
      <td height="26" colspan="2" background="backimg/bg1.gif"> <b><font color="#FFFFFF">��Ա�һ�����</font></b> 
        �ڶ��� (<strong>���������</strong>)</td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=4>�û���:</td>
      <td><%=rs("user")%></td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=4>�� ��:</td>
      <td><%=rs("quesion")%></td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=9>�� ��: </td>
      <td> <input name=answer type=text id=answer size=30> <input name="user" type="hidden" id="action" value="<%=rs("user")%>"> 
        <input name="action" type="hidden" id="action" value="step3"> </td>
    </tr>
    <%
	else
	response.Write("<tr><td colspan=2 bgcolor=#FFFFFF>�����û�������!û���ҵ����û�</td></tr>")
	end if
	case "step3"
	set rs=conn.execute("select quesion,answer,user from userinfo where user='"&request.form("user")&"'")
	if (not rs.eof) and (not rs.bof) then
	  if rs("answer")=MD5(trim(request.Form("answer"))) then
	%>
	    <tr bgcolor=#FFFFFF> 
      <td height="26" colspan="2" background="backimg/bg1.gif"> <b><font color="#FFFFFF">��Ա�һ�����</font></b> 
        ������ (<strong>�޸�����</strong>)</td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=4>�û���:</td>
      <td><%=rs("user")%> <input name="user" type="hidden" id="action" value="<%=rs("user")%>"> 
        <input name="action" type="hidden" id="action" value="step4"> <input name="answer" type="hidden" id="answer" value="<%=request.Form("answer")%>"> 
      </td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td height=1>�� ��:</td>
      <td><%=rs("quesion")%></td>
    </tr>
    <tr bgcolor=#FFFFFF>
      <td height=2>�� ��:</td>
      <td><%=request.Form("answer")%></td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td height=9>�� �룺</td>
      <td><input name=password1 type=password id=password12 size=30></td>
    </tr>
    <tr bgcolor=#FFFFFF> 
      <td width=17% height=9>ȷ������: </td>
      <td> <input name=password2 type=password id=password1 size=30></td>
    </tr>
    <%
  else
  response.Write("<tr><td colspan=2 bgcolor=#FFFFFF>��������𰸲���ȷ!����������!</td></tr>")
  end if
else
response.Write("<tr><td colspan=2 bgcolor=#FFFFFF>�����û�������!û���ҵ����û�</td></tr>")
end if
case "step4"
%>
	  <tr bgcolor=#FFFFFF> 
      <td height="26" colspan="2" background="backimg/bg1.gif"> <b><font color="#FFFFFF">��Ա�һ�����</font></b> 
        <strong>(ȡ������ɹ�)</strong></td>
    </tr><tr><td height=10 colspan=2 bgcolor=#FFFFFF><%=errmess%><br>
       <li>���ѳɹ����޸�������.</td>
    </tr>
<%
end select
%>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td height="10" colspan="2">
	  <%if action<>"step4" then%>
	  <input type="submit" name="Submit" value="��һ��"> 
	  <%else%>
	   <input type="button" name="Submit" onClick="location='Login.asp'"value="�� ��"> 
	 <%end if%>
      </td>
    </tr>
  </table>
</form>
<!--#include file="end.asp" -->
</body>
</html>
