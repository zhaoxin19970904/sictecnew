<!--#include file="CONN.ASP" -->
<%
dim rs,sql,name,ly,czlb,id,rid,lb
if not userislogin then
founderr=true
errmess=errmess&error1
end if
if userloginlock=1 then
founderr=true
errmess=errmess&error2
end if
'���������ӳ���
sub addIntoData()
dim czlb,id,rid,lysting
id=request.Form("id")
rid=request.Form("rid")
czlb=request.Form("czlb")
if id<>"" and czlb<>"addtz" then
call checkiflocked(id)
end if
if founderr then
exit sub
end if
set rs=server.createobject("adodb.recordset")
select case czlb
case "addtz"
call updateftuser(true)
sql="select * from borecorder"
rs.open sql,conn,1,3
rs.addnew
rs("retime")=now()
lysting=request.form("ly")
case "edittz"
sql="select * from borecorder where id="&id
rs.open sql,conn,1,3
lysting=request.form("ly")&vbCrlf&vbCrlf&" [color=#FF00FF]�����ѱ��޸Ĺ�,�޸�ʱ��Ϊ:[/color]"&now()
case "retz"
call updateftuser(false)
sql="select * from rely"
rs.open sql,conn,1,3
rs.addnew
rs("rid")=id
lysting=request.form("ly")
case "editretz"
sql="select * from rely where id="&rid
rs.open sql,conn,1,3
lysting=(request.form("ly")&vbCrlf&vbCrlf&" [color=#FF00FF]�����ѱ��޸Ĺ�,�޸�ʱ��Ϊ:[/color]"&now())
case "vote"
call updateftuser(false)
sql="select * from rely"
rs.open sql,conn,1,3
rs.addnew
lysting=request.form("ly")
rs("rid")=id
case "revote"
call updateftuser(false)
sql="select * from rely"
rs.open sql,conn,1,3
rs.addnew
lysting=request.form("ly")
rs("rid")=id
case "editvote"
sql="select * from rely where id="&rid
rs.open sql,conn,1,3
lysting=request.form("ly")&vbCrlf&vbCrlf&" [color=#FF00FF]�����ѱ��޸Ĺ�,�޸�ʱ��Ϊ:[/color]"&now()
end select
rs("name")=request.form("name")
rs("heat")=request.form("heat")
rs("ly")=lysting
rs("lb")=request.form("lb")
rs("user")=loginuser
rs("time")=now()
rs("ip")=Request.ServerVariables("REMOTE_ADDR")
if id="" then 
id=rs("id")
end if
rs.update
founderr=true
errmess=errmess&"<li>��������Ѿ��ɹ��ķ�����̳��,3 �����Զ������������������."
errmess=errmess&"<script>window.tm = setInterval(""location.href='type.asp?id="&id&"'"", 3000)</script>"
errmess=errmess&"<li> <li>"&thispage_name(request.ServerVariables("SCRIPT_NAME"),"url")
errmess=errmess&"<li> <li><a href=type.asp?id="&id&">�������������</a>"
end sub

if request.form("sublit")<>"" then
addIntoData()
else
czlb=request.QueryString("czlb")
lb=request.QueryString("lb")
id=request.QueryString("id")
rid=request.QueryString("rid")
if czlb="" or lb="" then
founderr=true
errmess=errmess+"<li>�Բ����ֲ�������,�������������û�д���ȷ�����ӽ���.���<a href=index.asp>��̳</a>"
end if
if (czlb<>"addtz") and id="" then
founderr=true
errmess=errmess+"<li>�Բ���!���ֲ�������,�������������û�д���ȷ�����ӽ���.���<a href=index.asp>��̳</a>"
end if
if (czlb="editretz" or czlb="editvote" or czlb="revote") and rid="" then
founderr=true
errmess=errmess+"<li>�Բ���!���ֲ�������,�������������û�д���ȷ�����ӽ���.���<a href=index.asp>��̳</a>"
end if
select case czlb
case "addtz"
name="" 
ly=""
case "edittz"
set rs=conn.execute("select name,ly from borecorder where id="&id&" and user='"&loginuser&"'")
if not rs.eof then
name=rs("name")
ly=rs("ly")
else
founderr=true
errmess=errmess+"<li>�Բ���!�޷���ʾ��Ҫ�༭������.<li>ԭ��: �����û�б༭�����ӵ�Ȩ����ע�����ȷ���� <a href=index.asp>����</a>"
end if

case "retz"
set rs=conn.execute("select name,lock from borecorder where id="&id)
if not rs.eof then
name="[�ظ�:]"&rs("name")
else
founderr=true
errmess=errmess+"<li>�Բ���!û���ҵ���ظ���ָ�������.�����ȷ����<a href=index.asp>����</a>"
end if

case "editretz"
set rs=conn.execute("select name,ly from rely where id="&rid&" and user='"&loginuser&"'")
if not rs.eof then
name=rs("name")
ly=rs("ly")
else
founderr=true
errmess=errmess+"<li>�Բ���!�޷���ʾ��Ҫ�༭������.<li>ԭ��: �����û�б༭�����ӵ�Ȩ��.��ע�����ȷ���� <a href=index.asp>����</a>"
end if

case "vote"
set rs=conn.execute("select name,user,time,ly,lock from borecorder where id="&id)
if not rs.eof then
name="[����]:"&rs("name")
ly="[quote][color=#FF1493]���������� "&rs("user")&" �� "&rs("time")&"  �ķ��ԣ�[/color]"&vbCrlf&rs("ly")&"[/quote]"
else
founderr=true
errmess=errmess+"<li>�Բ���!û���ҵ���������ָ�������.�����ȷ����<a href=index.asp>����</a>"
end if

case "editvote"
set rs=conn.execute("select name,ly from rely where id="&rid&" and user='"&loginuser&"'")
if not rs.eof then
name=rs("name")
ly=rs("ly")
else
founderr=true
errmess=errmess+"<li>�Բ���!�޷���ʾ��Ҫ�༭������.<li>ԭ��: �����û�б༭�����ӵ�Ȩ��.��ע�����ȷ����<a href=index.asp>����</a>"
end if
case "revote"
set rs=conn.execute("select name,user,time,ly from rely where id="&rid)
if not rs.eof then
name="[����]:"&rs("name")
ly="[quote][color=#FF1493]���������� "&rs("user")&" �� "&rs("time")&"  �ķ��ԣ�[/color]"&vbCrlf&rs("ly")&"[/quote]"
else
founderr=true
errmess=errmess+"<li>�Բ���!û���ҵ���������ָ�������.�����ȷ����<a href=index.asp>����</a>"
end if
end select
end if
'��������Ƿ�����
sub checkiflocked(lockid)
dim rs
set rs=conn.execute("select lock from borecorder where id="&lockid)
if rs("lock")="locked" then
founderr=true
messerr=messerr&"<li>�Բ���!�������ѱ�����Ա����,���ٽ��ܻظ�.<li><a href=type.asp?id="&lockid&">��������</a>"
end if
end sub
%>
<!--#include file="mymem.asp" -->
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--

function form1_onsubmit() 
{
  if(document.FORM1.name.value.length<1)
 {
   alert("���������������");
   document.FORM1.name.focus();
   return false;
 }
   if(document.FORM1.ly.value.length<1)
 {
   alert("����д����Ҫ���������");
   document.FORM1.ly.focus();
   return false;
  }
}
function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="sublit"||tempobj.Reset.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
//-->
</SCRIPT>

<FORM name=FORM1 onSubmit="return form1_onsubmit()" action=AddIn.asp?lb=<%=lb%>&czlb=ok method=post>
  <table width="898" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="787" height="19"> </td>
    </tr>
    <tr>
      <td> 
<%
if founderr then
founderror(errmess)
end if
call tabletop(true,"top")
%>
        <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
          <tr bgcolor="#92b9fb" background="backimg/bg1.gif"> 
            <td height="27" colspan="2" background="backimg/bg1.gif"> <input name="lb" type="hidden" id="lb" value="<%=lb%>"> 
              <input name="czlb" type="hidden" id="czlb" value="<%=czlb%>"> 
              <input name="id" type="hidden" id="id" value="<%=id%>"> 
              <input name="rid" type="hidden" id="rid" value="<%=rid%>"> 
              <font color="#FFFFFF"><strong>��̳������ </strong></font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%"><b>�������</b><SELECT name=fonta onchange="document.FORM1.name.value+=fonta.value">
              <OPTION selected value="">ѡ����</OPTION> 
			  <OPTION value=[ԭ��]>[ԭ��]</OPTION> 
              <OPTION value=[ת��]>[ת��]</OPTION> 
			  <OPTION value=[��ˮ]>[��ˮ]</OPTION> 
              <OPTION value=[����]>[����]</OPTION> 
			  <OPTION value=[����]>[����]</OPTION> 
              <OPTION value=[�Ƽ�]>[�Ƽ�]</OPTION> 
			  <OPTION value=[����]>[����]</OPTION> 
              <OPTION value=[ע��]>[ע��]</OPTION> 
			  <OPTION value=[��ͼ]>[��ͼ]</OPTION>
              <OPTION value=[����]>[����]</OPTION> 
			  <OPTION value=[����]>[����]</OPTION>
              <OPTION value=[����]>[����]</OPTION>
			  </SELECT>
			  </td>
            <td width="81%"> <input name="name" type="text" id="tzname" value="<%=name%>" size="60" maxlength="100" > 
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%"><strong>�ļ��ϴ�</strong> <a href="#" title="gif<br>jpg<br>bmp<br>jpeg<br>png">����</a> <br> </td>
            <td bgcolor="#FFFFFF"><IFRAME name=open src="upfilex.asp?lb=<%=lb%>" frameBorder=0 width="610"  height="30"></IFRAME></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%"> <strong>��������</strong>: </td>
            <td width="81%" rowspan="2"> <!--#include file="getubb.asp" --> <textarea name="ly" rows="12" wrap="file" cols="75" title="��Ctrl+Enter���ύ����" onkeydown=ctlent()><%=ly%></textarea> 
              <br> 
<%
call listpicimg()
sub listpicimg()
dim i
for i=1 to 28
	if len(i)=1 then i="0" & i
	response.write "<img src=""pic/em"&i&".gif"" border=0 onclick=""insertsmilie('[em"&i&"]')"" style=""CURSOR: hand"">&nbsp;"
if i=14 then
response.write"<br>"
end if
next
end sub
%> </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" height="231"> <strong>����</strong><br>
<%
if cint(bbs_seting(2))=1 then
response.Write("HTML����:��֧��")
else
response.Write("HTML����:֧��")
end if
if cint(bbs_seting(3))=0 then
response.Write("<br>UBB����:��֧��")
else
response.Write("<br><a href=helpubb.asp>UBB</a>����:֧��")
end if
%></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2" align="center"> 
              <input name=sublit type=submit id="sublit" value="�� ��"> &nbsp;&nbsp; 
              <input type="reset" tolowercase name="Reset" value="�� ��"> </td>
          </tr>
        </table>
        <%call tabletop(true,"down")%>
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
<script language=javascript>
ie = (document.all)? true:false
if (ie){
function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.FORM1.submit();}}
}
</script>
</FORM>
<!--#include file="end.asp" -->
</body>
</html>
