<!--#include file="CONN.ASP" -->
<!--#include file="mymem.asp" -->
<table width="898" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="772"> 
      <form name="form1" method="get" action="">
  �û���: 
  <input type="text" name="suser" size="10">
  <select name="sel">
    <option>����</option>
    <option value="jibie">������</option>
    <option value="time">ע��ʱ��</option>
    <option value="userjf">����</option>
    <option value="atime">����ʱ��</option>
    <option value="ftz">������</option>
   <option value="kill">ID���</option>
  </select>
  <input type="submit" name="Submit" value="����/����">
</form>
</td>
  </tr>
</table>
<%
call tabletop(true,"top")
dim suser,sel,rs,sqltext,i
suser=request.querystring("suser")
sel=request.querystring("sel")
set rs=server.createobject("adodb.recordset")
if  suser<>""  then 
sqltext="select * from userinfo where  user = '"&suser&" ' "
elseif  sel<>"" then
select case sel
case"jibie"
sqltext="select * from userinfo order by jibie desc "
case"time"
sqltext="select * from userinfo order by time desc "
case"userjf"
sqltext="select * from userinfo order by userjf desc "
case"atime"
sqltext="select * from userinfo order by atime desc "
case"ftz"
sqltext="select * from userinfo order by ftz desc "
case"kill"
sqltext="select * from userinfo order by kill desc ,id desc"
end select
else
sqltext="select * from userinfo"
end if
rs.open sqltext,conn,1,1

dim MaxPerPage
MaxPerPage=50

dim text,checkpage,CurrentPage
text="0123456789"
 Rs.PageSize=MaxPerPage
for i=1 to len(request("page"))
   checkpage=instr(1,text,mid(request("page"),i,1))
   if checkpage=0 then
      exit for 
   end if
next

If checkpage<>0 then
      If NOT IsEmpty(request("page")) Then
        CurrentPage=Cint(request("page"))
        If CurrentPage < 1 Then CurrentPage = 1
        If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
      Else
        CurrentPage= 1
      End If
      If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
   CurrentPage=1
End if

if (not rs.eof) or not (rs.bof) then
call list()
call showpages("&suser="&suser&"&sel="&sel)
else
call list()
end if
'��ʾ���ӵ��ӳ���
Sub list()%>
<table width="898" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
  <tr> 
    <td height="27" colspan="2" align="center" background="backimg/bg1.gif"> �û���</td>
    <td height="27" align="center" background="backimg/bg1.gif">����</td>
    <td height="27" align="center" background="backimg/bg1.gif">�Ա�</td>
    <td height="27" align="center" background="backimg/bg1.gif">email</td>
    <td height="27" align="center" background="backimg/bg1.gif"> ����</td>
    <td height="27" align="center" background="backimg/bg1.gif"> ����ʱ��</td>
    <td height="27" align="center" background="backimg/bg1.gif">������</td>
    <td height="27" align="center" background="backimg/bg1.gif">��������</td>
    <td height="27" align="center" background="backimg/bg1.gif">ע��ʱ��</td>
    <td height="27" align="center"  background="backimg/bg1.gif" title="����û�ID�������û�ID���(��Աר��)">��ID</td>
  </tr>
  <%
if not rs.eof then
i=1
  do while not rs.eof 
%>
  <tr bgcolor="#FFFFFF"> 
    <td> <a href="usertype.asp?user=<%=rs("user")%>" title=������û�������Ϣ ><%=rs("user")%></a></td>
    <td align="center"><a href="#"><img src="pic/m_news.gif" width="21" height="14" border="0" onClick="MM_openBrWindow('sendmess.asp?user=<%=rs("user")%>','','scrollbars=yes,resizable=yes,width=400,height=350')" alt="�� <%=rs("user")%> ��������Ϣ"></a></td>
    <td align="center"><%=rs("jibie")%> </td>
    <td align="center"> <%=rs("xb")%> </td>
    <td> 
      <%if rs("setuser1")=1 then response.write rs("email") else response.write"����" end if%>
    </td>
    <td align="center"> <%=rs("userjf")%> </td>
    <td align="center"> <%=rs("atime")%> ��</td>
    <td align="center"> <%=rs("ftz")%> </td>
    <td align="center"> <%=rs("jhtz")%> </td>
    <td align="center"> <%=rs("time")%> </td>
    <td align="center"> 
     <input name="kill" type="checkbox" onClick=location.href='caozuo.asp?czlb=��ID&id=<%=rs("id")%>'  <% if rs("kill")=1 then response.write"checked title=���û�ID�ѱ�����Ա���" end if%> > 
	 <%
	  if master then 
	 response.Write("<a href=admin_cz.asp?czlb=deluser&id="&rs("id")&" onclick=""{if(confirm('ɾ�����û���ȷ��ɾ����?')){return true;}return false;}"">ɾ</a>") 
	 end if 
	 %></td>
  </tr>
  <%
   if i >= MaxPerpage then exit do
   rs.movenext
   i=i+1
   loop
else
response.Write("<tr><td bgcolor=#FFFFFF colspan=11>Ŀǰû���κ�ע���û�</td></tr>")
end if
%>
</table>
<%
End sub
%>
<%
'��׼��ҳ�ӳ���URLparameterΪURL������
sub showpages(URLparameter)
dim pages,soonhost,pagename,nx,x,ISselect,tenpages
pages=cint(request.QueryString("page"))
if pages=0 then
pages=1
end if
if pages>rs.pagecount then
pages=rs.pagecount
end if
pagename=request.ServerVariables("SCRIPT_NAME")

response.Write("<table width=898 border=0 align=center cellpadding=2 cellspacing=1 bgcolor=#92b9fb>")
response.Write(" <tr bgcolor=#FFFFFF>")
response.Write("  <td width=734 class=main1>")

tenpages=cint(pages/10-0.5001)
nx=tenpages*10
response.Write("ҳ��:<b>"&pages&"/"&rs.pagecount&"</b>&nbsp;ÿҳ<b>"&MaxPerPage&"</b>&nbsp;������<b>"&rs.recordcount&"</b>&nbsp;")
if request("page")>1 then
response.write "&nbsp;<a href='"&pagename&"?page=1"&URLparameter&"' title=��ҳ><font face=webdings>9</font></a>"
else
response.write "&nbsp;<font face=webdings>9</font>"
end if

if nx>=10 then
response.write "&nbsp;<a href='"&pagename&"?page="&(nx-10)&URLparameter&"' title=��ʮҳ><font face=webdings>7</font></a>"
else
response.write "&nbsp;<font face=webdings>7</font>"
end if

for x=1 to 10
nx=nx+1
if pages=nx then
response.write ("<b>&nbsp;<font color=red>"&nx&"</font></b>")
else
Response.write "<b>&nbsp;<a href='"&pagename&"?page="&nx&URLparameter&"'>"&nx&"</a></b>"
end if
if nx>=rs.pagecount then exit for
next

if rs.pagecount>=10 and nx<rs.pagecount then
response.write "&nbsp;<a href='"&pagename&"?page="&(nx+1)&URLparameter&"' title=��ʮҳ><font face=webdings>8</font></a>"
else
response.write "&nbsp;<font face=webdings>8</font>"
end if
if pages<rs.pagecount then
response.write "&nbsp;<a href='"&pagename&"?page="&rs.pagecount&URLparameter&"' title=βҳ><font face=webdings>:</font></a>"
else
response.write "&nbsp;<font face=webdings>:</font>"
end if

response.Write("</td><td width=""25"" class=main1 >")

response.Write("<select name=pagd onChange=setTimeout(location.href=this.value,3000)>")
soonhost=1
do while not soonhost = (rs.PageCount+1) 
if (soonhost)=pages then
ISselect=" selected"
else
ISselect=""
end if
response.Write("<option value="&pagename&"?Page="&soonhost&URLparameter&ISselect&">"&soonhost&"</option>")
soonhost=soonhost+1
loop

response.Write("</select> </td></tr></table>")
end sub
call tabletop(true,"down")
%>
<!--#include file="end.asp" -->
</body>
</html>
