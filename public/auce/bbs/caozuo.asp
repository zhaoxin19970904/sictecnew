<!--#include file="conn.asp"-->
<%
if not master then 
response.cookies("rewin")("errmess")="<li><b>��ҳ��Ϊ����ҳ��.</b><li>ԭ��:<li>����ܲ��߱������ҳ���Ȩ��."
response.Redirect("error.asp?founderr=mess")
end if

if len(request.form("FORM1")) then
call intoadmincz()
end if

sub intoadmincz()
'on error resume next
dim czyy,czname,czid,atlt,bczuser,lb,rid,czlb,rs
czyy=request.form("czyy")
czname=request.Form("czname")
czid=request.form("czid")
atlt=request.form("atlt")
bczuser=request.form("bczuser")
lb=request.form("lb")
rid=request.Form("rid")
czlb=request.form("czlb")
if  czyy="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('����д�� ����ԭ��');history.back()</script>"
Response.end
end if
dim czzl,fenl,i,fenz,nex,fen,adddatesm
czzl=array("�ö�","ɾ��","����","����","ѡ��","����","���","����","����","�ƶ�","���","ɾ��","���","��ID","��ID")
fenl=array(0,-10,-50,100,300,0,0,0,0,0,0,0,0,0,0)
for i=0 to ubound(czzl)
if czzl(i)=czlb then
fenz=fenl(i)
end if
next
fen=fenz&fen
fen=czyy&":  ��ֵ����: "&fen
errmess="<li>"
adddatesm=(czname&" -- "&fen)
nex="<br>&nbsp;& -- "&czname&"<li><li><a href=list.asp?lb="&lb&" >���������б�</a><li><li><a href=type.asp?id="&czid&">����������ʾ</a>"
conn.execute("insert into bbs_log([czid],[lbname],[lb],[rid],[ip],[bczuser],[czuser],[czlb],[czlr],[time]) values ('"&czid&"','"&atlt&"','"&lb&"','"&rid&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&bczuser&"','"&loginuser&"','"&czlb&"','"&adddatesm&"','"&now()&"')")
select case czlb
case "�ö�"
set rs=conn.execute("select [top] from borecorder order by top desc")
conn.execute("update borecorder set [top]="&(rs("top")+1)&" where id="&czid)
errmess=errmess&czlb&rs("top")&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "ȥ��"
conn.execute("update borecorder set [top]=0 where id="&czid)
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "ɾ��"  
conn.execute("delete from borecorder where id="&czid)
conn.execute("update userinfo set deltz=deltz+1,userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "����"
conn.execute("update userinfo set userjf=userjf+"&fen&" where user='"&bczuser&"'")
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "����"
conn.execute("update borecorder set jh=911 where id="&czid)
conn.execute("update user set userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "����"
conn.execute("update borecorder set jh=8 where id="&czid)
conn.execute("update userinfo set jhtz=jhtz+1,userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "ȥ��"
conn.execute("update borecorder set jh=0 where id="&czid)
conn.execute("update userinfo set jhtz=jhtz+1,userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "����"
conn.execute("update borecorder set lock='locked' where id="&czid)
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "����"
conn.execute("update borecorder set lock='' where id="&czid)
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "����"
conn.execute("update borecorder set lb='"&lb&"' where id="&czid)
errmess=errmess&czlb&" �� "&zlbltname(lb)&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "����"
conn.execute("update borecorder set kill=1 where id="&czid)
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "���"
conn.execute("update borecorder set kill=0 where id="&czid)
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex
case "��ID"
conn.execute("update userinfo set kill=1 where id="&czid)
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: 0"
errmess=errmess&"<br>&nbsp;<p><a href=listalluser.asp>����</a> &nbsp;</p>"
case "��ID"
conn.execute("update userinfo set kill=0 where id="&czid)
errmess=errmess&czlb&" �����ɹ�!�ִ���ֵ����Ϊ: 0"
errmess=errmess&"<br>&nbsp;<p><a href=listalluser.asp>����</a> &nbsp;</p>"
case "���"
conn.execute("update rely set kill=1 where id="&rid)
errmess=errmess&czlb&" �����ɹ�!�ִ���ֵ����Ϊ: 0"
errmess=errmess&"<br>&nbsp;<p><a href=type.asp?id="&czid&">����</a> &nbsp;</p>"
case"���"
conn.execute("update rely set kill=0 where id="&rid)
errmess=errmess&czlb&" �����ɹ�!��ֵ����Ϊ: "&fen&nex
errmess=errmess&"<br>&nbsp;<p><a href=type.asp?id="&czid&">����</a> &nbsp;</p>"
case "ɾ��"
conn.execute("delete from rely where id="&rid)
conn.execute("update userinfo set deltz=deltz+1,userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" �����ɹ�! ��ֵ����Ϊ: "&fen&nex&" <a href=type.asp?id="&czid&">������ʾ����</a>"
case "ɾ��"
end select
end sub
%>
<!--#include file="mymem.asp" -->

  
<table width="898" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
  <tr> 
      <td height="31" background="backimg/bg1.gif"> <strong><font color="#FFFFFF">�߼�����Ա����</font></strong></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF">
<%
if errmess<>"" then 
response.write(errmess)
else
call loadingmess()
end if
sub loadingmess()
dim czlb,id,rs,sql,bczuser,lb,czname,atlt,rsr
czlb=request.querystring("czlb")
id=request.querystring("id")
lb=request.QueryString("lb")
if id<>"" and lb<>"" then
set rs=conn.execute("select user,id,lb,name,kill,lock,top,jh from borecorder where id="&id)
if rs.eof or rs.bof then
response.Write("û���ҵ�������")
response.End()
end if 
bczuser=rs("user")
atlt=zlbltname(rs("lb"))
end if

select case czlb
case "�ö�"
if rs("top")>0 then
czlb="ȥ��"
czname="ȥ�������ӵ��ö�"
else
czname="�����ӹ̶�����̳����"
end if
case "ɾ��"
czname="ɾ������"
case "����"
czname="�ƶ�����"
case "����"
if rs("kill")=0 then
czlb="����"
czname="�������"
else
czname="������"
czlb="���"
end if
case "����"
if rs("lock")="locked" then
czlb="����"
czname="�������"
else
czlb="����"
czname="��������"
end if
case "����"
czname="�������û�����"
case "����"
if rs("jh")=8 then
czlb="ȥ��"
czname="ȥ�������ھ�����"
else
czname="������Ϊ����"
end if
case "��ID"
   set rs=conn.execute("select user,kill from userinfo where id="&id)
   if rs("kill")=1 then 
   czlb="��ID"
   czname="������û�ID�ķ��"
   else
   czname="����û�ID"
   end if
   bczuser=rs("user")
case "���"
set rsr=conn.execute("select user,kill,lb from rely where id="&request.QueryString("rid"))
if rsr.bof or rsr.eof then 
response.Write("�޴�����,����ʧ��")
response.End()
end if
bczuser=rsr("user")
if rsr("kill")=0 then
czlb="���"
czname="��������ظ�"
else
czname="���������ظ�����"
czlb="���"
end if
case "ɾ��"
set rsr=conn.execute("select user,kill,lb from rely where id="&request.QueryString("rid"))
czname="ɾ���ظ�������"
bczuser=rsr("user")
end select
%>
<form name="FORM1" method="post" action="caozuo.asp?lb=<%=lb%>">
	  <table width="100%" border="0">
          <tr> 
            <td><strong>��������: <font color="#FF0000"><%=czname%></font></strong></td>
          </tr>
          <tr> 
            <td><input type="hidden" name="bczuser" value="<%=bczuser%>"> 
              <input name="atlt" type="hidden" id="atlt" value="<%=atlt%>"> 
              <input type="hidden" name="czname" value="<%=czname%>"> 
              <input type="hidden" name="czid" value="<%=id%>"> 
              <input name="czlb" type="hidden" id="czlb" value="<%=czlb%>"> 
              <input name="lb" type="hidden" id="lb" value="<%=lb%>"> 
              <input name="rid" type="hidden" id="rid" value="<%=request.QueryString("rid")%>">
              ����ԭ��: 
              <select name="kczyy" size=1 onChange="document.FORM1.czyy.value=kczyy.value">
                <option value="">�Զ���</option>
                <option value="��ˮ">��ˮ</option>
                <option value="���">���</option>
                <option value="����">����</option>
                <option value="������">������</option>
                <option value="���ݲ���">���ݲ���</option>
                <option value="�ظ�����">�ظ�����</option>
                <option value="����־ѡ��">����־ѡ��</option>
                <option value="ѡΪ����">ѡΪ����</option>
              </select> <input name="czyy" type="text" id="czyy" size="60" ></td>
          </tr>
          <tr> 
            <td> <% if request.querystring("czlb")="����" then%>
              ��ֵ: 
              <select name=fen>
                <option value="0" selected>0</option>
                <script language=JavaScript><!--               
          for(i=-2000;i<2001;i++) document.write('<option>'+i)               
            //-->
         </script>
              </select> <%end if%> <%
if request.querystring("czlb")="����" then
call ltnamelist()
end if
%> </td>
          </tr>
          <tr> 
            <td align="center"> <input name=FORM1 type=submit  value="�� ��"> &nbsp; 
              <input name="Submit" type="button" onClick="history.go(-1);" value="�� ��"> 
            </td>
          </tr>
        </table>
		</form>
		<%end sub%>
		</td>
    </tr>
  </table>

<%
sub ltnamelist()
dim rs
set rs=conn.execute("select ltlb,id from ltlb order by b_order desc")
response.write" &nbsp; ��ѡ��Ҫ�Ƶ�����̳&nbsp; "
response.write"<select name=""ltlb"">"
if not rs.eof then
do while not rs.eof
response.write"<option value="""&rs("id")&""">"&rs("ltlb")&"</option>" 
rs.movenext
loop
response.write"</select>"
end if
end sub
%>
<!--#include file="end.asp" -->
</body>
</html>
