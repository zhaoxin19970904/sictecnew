<!--#include file="admin_conn.asp" -->
<%
if request.Form("Submit")<>"" then
call ltmessetruyg() '������̳�������ͻظ���,����һ��ʱ�������ݽ��и���
call updatereguser() 'ע���û���
end if
sub ltmessetruyg()
response.Write("���ڶ������ݽ��и���....")
dim rs
	 set rs=conn.execute("select tzsum,retzsum,id from ltlb")
	 do while not rs.eof 
conn.execute("update ltlb set tzsum="&lastboardlyx(rs("id"),"tzs")&" where id="&rs("id"))
conn.execute("update ltlb set retzsum="&lastboardlyx(rs("id"),"tz1s")&" where id="&rs("id"))
	 rs.movenext
	 loop
response.Write("<br>��̳���ݸ������!")
end sub

function lastboardlyx(zlt,tzs)
dim rs,i,rtzs,sql
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select id,rely from borecorder where lb='"&zlt&"'order by id desc"
rs.open sql,conn,1,1
if tzs="tzs" then
lastboardlyx=rs.recordcount
else
for i=1 to rs.recordcount
rtzs=rtzs+rs("rely")
rs.movenext
next
lastboardlyx=rtzs
end if
end function

'call updaterelys() '����������
sub updaterelys()
dim rs,rs1,sql
set rs1=server.createobject("adodb.recordset")
set rs=conn.execute("select id from borecorder")
do while not rs.eof 
sql="select id from rely where rid='"&rs("id")&"'"
rs1.open sql,conn,1,1
if not rs1.eof then
conn.execute("update borecorder set [rely]="&rs1.recordcount&" where id="&rs("id"))
end if
rs.movenext
loop
end sub
sub updatereguser()
dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select user from userinfo"
rs.open sql,conn,1,1
conn.execute("update admin_copyc set reguser="&rs.RecordCount&" where id=1")
end sub
%>
<link href="DEFAULT.css" rel="stylesheet" type="text/css">

<form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>">
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
    <tr> 
      <td height="22">������̳����,���¼���������Ŀ,����һ��ʱ�����һ������</td>
    </tr>
    <tr>
      <td height="22"><input type="submit" name="Submit" value="ȷ������"></td>
    </tr>
  </table>
</form>