<!--#include file="CONN.ASP" -->
<%
dim lb
if not userislogin then
founderr=true
errmess=errmess&error1
end if
if userloginlock=1 then
founderr=true
errmess=errmess&error2
end if
lb=request.QueryString("lb")
if len(request.form("FORM1")) then
call addtpdate()
end if
sub addtpdate()
if founderr then
exit sub
end if
dim rs,sqltext,tply,ltsel,x,i,bz,bzc,from_address,tpid,id
x=0
tply=replace(trim(request.Form("tply")),vbCrlf,"|")
		ltsel=split(tply,"|")
		for i=0 to ubound(ltsel)
        if ltsel(i)<>"" then
        x=x+1		
        bz=ltsel(i)+"|"+bz
        bzc="0|"+bzc
        end if
        next
if x<2 or x>10 then
founderr=true
errmess=errmess&"<li>��������:<li>ԭ��:ͶƱ��Ŀ��������2��ʹ���10���!"
exit sub
end if 
if request.Form("name")="" or request.Form("ly")="" then
founderr=true
errmess=errmess&"<li>����:������������ݶ�����,��ע�������"
exit sub
end if
call updateftuser(true)
set rs=server.createobject("adodb.recordset")
rs.open "lttp",conn,1,3
rs.Addnew
rs("tply")=bz
rs("count")=bzc
rs("name")=request.form("name")
rs("datetime")=now()
rs("tplb")=request.form("tplb")
rs("timeout")=request.form("timeout")
tpid=rs("id")
rs.update

set rs=server.createobject("adodb.recordset")
rs.open "borecorder",conn,1,3

'���һ����¼�����ݿ�
rs.addnew
rs("istp")=tpid
rs("name")=request.form("name")
rs("heat")=request.form("heat")
rs("ly")=request.form("ly")
rs("lb")=request.form("lb")
rs("user")=loginuser
rs("time")=now()
rs("retime")=now()
rs("ip")=Request.ServerVariables("REMOTE_ADDR")
id=rs("id")
rs.update
founderr=true
errmess=errmess&"<li>��������Ѿ��ɹ��ķ�����̳��,3 �����Զ������������������."
errmess=errmess&"<script>window.tm = setInterval(""location.href='type.asp?id="&id&"'"", 3000)</script>"
errmess=errmess&"<li> <li>"&thispage_name(request.ServerVariables("SCRIPT_NAME"),"url")
errmess=errmess&"<li> <li><a href=type.asp?id="&id&">�������������</a>"
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
   document.form1.name.focus();
   return false;
 }
  if(document.FORM1.tply.value.length<1)
 {
   alert("������ͶƱ��Ŀ.");
   document.form1.name.focus();
   return false;
 }
   if(document.FORM1.ly.value.length<1)
 {
   alert("����д����Ҫ���������");
   document.FORM1.ly.focus();
   return false;
  }
}
//-->
</SCRIPT>
<%
if founderr=true then
call founderror(errmess)
end if
%>
<FORM action=<%=request.ServerVariables("SCRIPT_NAME")%>?lb=<%=lb%> method=post name=FORM1 id="form1"   onsubmit="return form1_onsubmit()" language=javascript>
  <table width="898" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
    <tr bgcolor="#6699CC"> 
      <td height="27" colspan="3" background="backimg/bg1.gif"> <font color="#FFFFFF"><strong>��̳����ͶƱ</strong></font><b> 
        <input name="lb" type="hidden" id="lb" value="<%=lb%>">
        </b></td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td width="148"> <b>�������</b> <SELECT name=fonta onchange="document.FORM1.name.value+=fonta.value">
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
        </SELECT> </td>
      <td width="363"> <input type="text" name="name" size="60" maxlength="100" > 
      </td>
      <td width="240">��Ч 
        <select name="timeout" >
          <option selected>����</option>
          <option value="10">10</option>
          <option value="15">15</option>
          <option value="30">30</option>
          <option value="60">60</option>
          <option value="90">90</option>
        </select>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td width="148"> <p><strong>ͶƱ��Ŀ</strong><br>
          ÿһ����Ϊһ����Ŀ<br>
          ����Ŀ�����������������10��<br>
          (����ʹ��HTML��ǩ��UBB��ǩ)</p></td>
      <td colspan="2"> <textarea name="tply" cols="60" rows="6" ></textarea> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE"""">
      <td width="148" height="30"><strong>�ļ��ϴ�</strong> <a href="#" title="gif,jpg,bmp,jpeg,png">����</a></td>
      <td colspan="2"><IFRAME name=open src="upfilex.asp?lb=<%=lb%>" frameBorder=0 width="610"  height="30"></IFRAME></td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td width="148" valign="top"> <strong>��������:</strong></td>
      <td colspan="2" rowspan="2"> <!--#include file="getubb.asp" --> <p> 
          <textarea name="ly" rows="8" wrap="file"  class=cid onKeyDown=ctlent() cols="70"></textarea>
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
%>
        </p></td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td width="148"> <strong><a href="helpubb.asp">����</a></strong><br>
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
%>
        <br> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td colspan="3"> <div align="center"> 
          <p><font color="#0000FF"> 
            <input name=FORM1 type=submit  value="�� ��">
            &nbsp;&nbsp; 
            <input type="reset" toLowerCase name="Reset" value="�� ��">
            </font></p>
        </div></td>
    </tr>
  </table>
  </center>
  </div>
</FORM>
<!--#include file="end.asp" -->
</body>
</html>
