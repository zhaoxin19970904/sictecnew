 <!--#include file="admin_conn.asp"-->
<%
dim id,user
id=request.form("id")
user=request.form("user")
if request.form("adduser")<>"" then
call addadminuser()
end if
if request.Form("deluser")<>"" then
call deladminuser()
end if
sub addadminuser()
dim jibie,rs,sqltext,ltbz
jibie=request.form("jibie")
set rs=server.createobject("adodb.recordset")
sqltext="select jibie,qx from userinfo where user='"&user&"'"
rs.open sqltext,conn,1,3
if rs.eof then
response.write"û�ҵ�������.��Ӳ��ɹ�<a href=# onClick='history.back()'> ����</a>"&user
exit sub
end if
rs("jibie")=jibie
rs("qx")=(request("id")&"|"&rs("qx"))
rs.update
set rs=server.createobject("adodb.recordset")
sqltext="select ltbz from ltlb where id="&request.Form("id")
rs.open sqltext,conn,1,3
ltbz=rs("ltbz")&user&"|"
rs("ltbz")=ltbz
rs.update
response.write "<script> parent.document.all.ltbz.innerHTML="""&rs("ltbz")&"""</script>"
response.write"���ӹ���Ա�ɹ�!<a href=# onClick='history.back()'> ����</a>"
response.end
end sub

sub deladminuser()
dim rs,sqltext,bz,i,ltsel
conn.execute("update userinfo set jibie='��ͨ�û�' where user='"&user&"'")
set rs=server.createobject("adodb.recordset")
sqltext="select ltbz,id from ltlb where id="&id
rs.open sqltext,conn,1,3
bz=""
		ltsel=split(rs("ltbz"),"|")
		for i=0 to ubound(ltsel)
		if ltsel(i)=user then
		ltsel(i)=""
		end if
if ltsel(i)<>"" then		
bz=ltsel(i)+"|"+bz
end if
next
rs("ltbz")=bz
rs.update
response.write "<script> parent.document.all.ltbz.innerHTML="""&rs("ltbz")&"""</script>"
response.write"ɾ������Ա�ɹ�!<a href=# onClick='history.back()'> ����</a>"
response.end
end sub
%>
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
<body leftmargin="0" topmargin="0">
<form name="form1" method="post" action="adadmin.asp">
  <input type="hidden" name="id" value="<%=request.QueryString("id")%>">
  <input type="text" name="user">
  <select name="jibie">
    <option value="����Ա">����Ա</option>
    <option value="�ܰ���">�ܰ���</option>
    <option value="վ��">վ��</option>
    <option value="����">����</option>
    <option value="��ί">��ί</option>
    <option value="�༭">�༭</option>
    <option value="׫����">׫����</option>
  </select>
  <input name="adduser" type="submit" id="adduser" value="���">
  <input name="deluser" type="submit" id="deluser" value="ɾ��">
</form>


