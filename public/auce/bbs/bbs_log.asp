<!--#include file="CONN.ASP" -->
<!--#include file="mymem.asp" -->
<%
if not master then 
response.cookies("rewin")("errmess")="<li><b>��ҳ��Ϊ����ҳ��.</b><li>ԭ��:<li>����ܲ��߱������ҳ���Ȩ��."
response.Redirect("error.asp?founderr=mess")
end if
dim rs,sql,MaxPerPage,page,x,z,nn,y,i,CurrentPage
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from bbs_log order by id desc"
rs.open sql,conn,1,1

MaxPerPage=100
if request("page")<>"" then
page=request("page")
else 
page=1
end if
dim text,checkpage
text="0123456789"
 Rs.PageSize=MaxPerPage
for i=1 to len(page)
   checkpage=instr(1,text,mid(page,i,1))
   if checkpage=0 then
      exit for 
   end if
next

If checkpage<>0 then
      If NOT IsEmpty(page) Then
        CurrentPage=Cint(page)
        If CurrentPage < 1 Then CurrentPage = 1
        If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
      Else
        CurrentPage= 1
      End If
      If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
   CurrentPage=1
End if
y=1
call lists()
if not rs.eof then
call showpages("&")
end if
sub lists()
response.write"<table width=770 border=0 cellpadding=3 cellspacing=1 align=center bgcolor=#92b9fb><tr bgcolor=#FFFFFF align=center> " 
response.write("<td background=backimg/bg1.gif height=27>���</td><td background=backimg/bg1.gif>��̳����</td>  <td background=backimg/bg1.gif>������ID</td> <td background=backimg/bg1.gif>��ID</td> <td background=backimg/bg1.gif>��������</td><td background=backimg/bg1.gif>��������</td>  <td background=backimg/bg1.gif>������</td> <td background=backimg/bg1.gif>����ʱ��</td><td background=backimg/bg1.gif>������IP</td><td background=backimg/bg1.gif>����ԭ��</td><td background=backimg/bg1.gif>��̳</td></tr>")
if rs.eof or rs.bof then
response.Write("<tr><td bgcolor=#FFFFFF colspan=11>��û�κ���־��¼</td></tr>")
end if
do while not rs.eof 
response.write"<tr align=center >"
for x=0 to 10
if x=9 then
nn="<a href=# title='"&rs(x)&"'>����</a>"
else
nn=rs(x)
end if
response.write "<td bgcolor=#FFFFFF>"&nn&"</td>"
next
response.write "</tr>"
rs.movenext
if y >= MaxPerpage then exit do
y=y+1
loop
response.write "</tr></table>"
end sub
%>
      </td>
    </tr>
  </table>
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

response.Write("<table width=770 border=0 align=center cellpadding=2 cellspacing=1 bgcolor=#92b9fb>")
response.Write(" <tr bgcolor=#FFFFFF>")
response.Write("<td width=734 class=main1>")

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
%>
 <!--#include file="end.asp" -->
</body>
</html>
