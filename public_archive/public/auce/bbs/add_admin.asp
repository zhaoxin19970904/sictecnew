 <!--#include file="admin_conn.asp"-->
<!--#include file="md5.asp" -->
<%
if request.Form("addadmin")<>"" then
call addadminuser()
end if
if request.Form("b1")<>"" then
call changepass()
end if
sub addadminuser()
dim user,psw,psw2,rs,sqltext
user=trim(replace(request.form("username"),"'",""))
psw=trim(replace(request.form("password1"),"'",""))
psw2=trim(replace(request.form("qqq"),"'",""))
set rs=conn.execute("select user from userinfo where user='"&user&"'")
if rs.eof or rs.bof then
errmess="<li>�߼��û�������ע���û���.���ӹ���Աʧ��!"
exit sub
else
conn.execute("update userinfo set jibie='����Ա' where user='"&user&"'")
end if
set rs=server.createobject("adodb.recordset")
sqltext="select * from admin"
rs.open sqltext,conn,1,3
if user="" or psw="" then
errmess="<li>�û������붼����Ϊ��!���ӹ���Աʧ��!"
exit sub
end if
if psw<>psw2 then 
errmess="<li>�����������벻һ��!���ӹ���Աʧ��!"
exit sub
end if
rs.addnew

rs("username")=user
rs("password")=MD5(psw)
rs.update
errmess="<li>�¹���Ա��ӳɹ�!"
end sub 

sub changepass() '���ĸ߼�����Ա����
dim username,password,password2,password3,rs,sql
if  request.form("username")<>""  then
username=trim(replace(request("username"),"'",""))
password=trim(replace(request("password"),"'",""))
password2=trim(replace(request("password2"),"'",""))
password3=trim(replace(request("password3"),"'",""))
if  password2="" or password3="" then
errmess="<li>�����벻��Ϊ��!�޸�ʧ��!"
exit sub
end if
If password2 <>password3  Then
errmess="<li>������������벻һ��!�޸�ʧ��!"
exit sub
End If
set rs=server.createobject("adodb.recordset")
Sql="Select password,username from admin where username='"&username&"'"
Rs.Open Sql,Conn,1,3
if not rs.eof then
If Rs("password")<>MD5(password) Then
errmess="<li>����������벻��ȷ!�޸�ʧ��!"
exit sub
End If
elseif  rs.eof  then
errmess="<li>�����ǹ���Ա�����޸�!�޸�ʧ��!"
exit sub
end if

Rs("password")=MD5(password2)
rs.update
end if
errmess="�����޸ĳɹ�!���������!"
end sub
%>
<html>
<head>
<title>��̳����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<div align="center"><br>
  <table width="450" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
    <%
call listadminuser()
sub listadminuser()
dim rs,sqltext,i
set rs=conn.execute("select username,id from admin")
do while not rs.eof
%>
    <tr bgcolor="#FFFFFF"> 
      <td width="198"> <%= rs("username")%> </td>
      <td width="167"><font color="#000000">�߼�����Ա</font> </td>
      <td width="71"> 
        <div align="center"><a href=admin_cz.asp?id=<%=rs("id")%>&czlb=deladm title="ɾ���˲�Ʒ���!(�벻Ҫ����ɾ�����)" onclick="{if(confirm('ȷʵҪ��������̳ɾ����ȷ��ɾ����?')){return true;}return false;}">ɾ�� 
          </a><a href="add_admin.asp?user=<%=rs("username")%>">�޸�</a></div>
      </td>
    </tr>
    <% 
rs.movenext
loop
end sub 
%>
  </table>
  
</div>
<form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>">
<%if request.QueryString("user")="" then%>
  <table width="450" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
    <tr><td bgcolor="#FFFFFF" colspan="2"><%=errmess%></td></tr>
	<tr bgcolor="#92b9fb"> 
      <td height="27" colspan="2" background="backimg/bg1.gif"> 
        <div align="center"><font size="4"><b><font color="#FF0000">�߼�����Ա���</font></b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="62" bgcolor="#FFFFFF">�û���</td>
      <td width="367" bgcolor="#FFFFFF" align="right"> 
        <input name="username" type="text"  size="30" maxlength="30">
      </td>
    </tr>
    <tr> 
      <td width="62" bgcolor="#FFFFFF">����</td>
      <td width="367" bgcolor="#FFFFFF" align="right"> 
        <input name="password1" type="password"  size="30" maxlength="30">
      </td>
    </tr>
    <tr> 
      <td width="62" bgcolor="#FFFFFF">�ظ�����</td>
      <td width="367" bgcolor="#FFFFFF" align="right"> 
        <input type="password" name="qqq" size="30" maxlength="30">
      </td>
    </tr>
    <tr> 
      <td width="62" bgcolor="#FFFFFF">��</td>
      <td width="367" bgcolor="#FFFFFF" align="right"> 
        <div align="center"> 
          <input name="addadmin" type="submit" id="addadmin" value="�� ��">
        </div>
      </td>
    </tr>
  </table>
  <%else%>
  <table width="450" border="0" align="center" cellspacing="1" bgcolor="#92b9fb">
    <tr> 
      <td align="center" bgcolor="#FFFFFF"><%=errmess%></td>
    </tr>
    <tr> 
      <td height="28" align="center" background="backimg/bg1.gif" bgcolor="#FFFFFF"><strong><font color="#FFFFFF">�߼�����Ա��������</font></strong></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#FFFFFF"> <table width=100% border=0 cellpadding=4 cellspacing=1>
          <tr bgcolor="#FFFFFF"> 
            <td width=23% class=link>����Ա���ƣ�</td>
            <td width=77%> <input name=username type=text class=link id="username" size=20 maxlength=20 value="<%=request.QueryString("user")%>"> 
              <font color="#FF3333">*</font> ��ĸ������</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width=23% height="31" class=link>�� �� �룺</td>
            <td width=77%> <input type=password name=password size=20 maxlength=20 class=link> 
              <font color="#FF3333">*</font> ��ĸ������</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width=23% class=link>�� �� �룺</td>
            <td width=77%> <input type=password name=password2 size=20 maxlength=20 class=link> 
              <font color="#FF3333">*</font> ��ĸ������</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width=23% class=link>ȷ�����룺</td>
            <td width=77%> <input type=password name=password3 size=20 maxlength=20 class=link> 
              <font color="#FF3333">*</font> ��ĸ������</td>
          </tr>
        </table>
        <br> <input type="submit" value="ȷ���޸�" name="b1">
        &nbsp;&nbsp; <input type="reset" name="Reset" value="������д"> <br>
        <br> </td>
    </tr>
  </table>
</form>  
<%end if%>
</body>
</html>

