<!--#include file="conn.asp"-->
<!--#include file="mymem.asp" -->
<!--#include file="ubbcode.asp" -->
<%
'=========================================================
' File: index.asp
' Version:1.0.1
' Date: 2003-8-23
' Script Written by cheng
'=========================================================
' Copyright (C) 2002-2003 mtvok.com. All rights reserved.
' Web: http://www.mtvok.com,http://www.dvking.cn,http://www.dvonlin.cn
' Email: zcl22@21cn.com
'=========================================================
dim bbsskin 'ȷ��������ʾ��ʽ,����,ƽ��
if request.QueryString("skin")<>"" then
 if request.QueryString("skin")=1 then
 response.Cookies("skinx")="tree"
 else
 response.Cookies("skinx")="nottree"
 end if
end if
if request.Cookies("skinx")="tree" then
bbsskin=1
else
bbsskin=0
end if
if request.Form("Submit")<>"" then
call adduserlttp()
end if
dim lb,id,rs,sqltext,name
id=request.QueryString("id")
conn.execute("update borecorder set click=click+1 where id="&id)
set rs=conn.execute("select * from borecorder where id="&id)
name=dvHTMLEncode(rs("name"))
lb=rs("lb")
id=rs("id")
%>
<table width="898" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="770"> <table width="100%" border="0">
        <tr> 
          <td width="55%"><a href="AddIn.asp?lb=<%=lb%>&czlb=addtz"><img src="pic/postnew.gif" width="72" height="21" border="0"></a> 
            <a href="tpaddIn.asp?lb=<%=lb%>"><img src="PIC/votenew.gif" width="72" height="21" border="0"></a> 
            <a href="AddIn.asp?id=<%=id%>&lb=<%=lb%>&czlb=retz"><img src="PIC/replybbs.gif" width="72" height="21" border="0"></a></td>
          <td width="24%">���Ǳ����ĵ� <b><%=rs("click")%></b> ���Ķ���</td>
          <td><%call chanagego(id,lb)%></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td> <%call tabletop(true,"top")%> <TABLE width=100% border=0 align="center" 
cellPadding=0 cellSpacing=0 borderColor=#92b9fb style="BORDER-COLLAPSE: collapse">
        <TBODY>
          <TR> 
            <TD width="91%" height=28 background="backimg/bg1.gif">&nbsp;<b><font color="#FFFFFF">* 
              <%=name%></font></b></TD>
            <TD width="9%" background="backimg/bg1.gif"><a href="foruser.asp?czlb=addf&id=<%=id%>"><img src="PIC/fav_add.gif" alt="�����ҵ��ղ�" width="16" height="16" border="0"></a> 
              &nbsp; <a href=# onClick="window.external.AddFavorite('http://<%=Request.ServerVariables("HTTP_HOST")&request.ServerVariables("SCRIPT_NAME")&"?id="&id%>', '<%=name%>')"><IMG SRC="pic/fav_add1.gif" BORDER=0 width=15 height=15 alt=�ѱ�������IE�ղؼ�></a> 
            </TD>
          </TR>
        </TBODY>
      </TABLE>
      <%
	  if rs("istp")<>0 then 
	  call typeusertp(rs("istp"))
	  end if
	  %> 
      <table width="100%" border="0"  align="center" cellpadding="0" cellspacing="1" bgcolor="#92b9fb"  class=tableborder1 style="table-layout; word-break:break-all"#92b9fb>
        <tr> 
          <td width="149" align="center" valign="top" bgcolor="#FFFFFF"><br>
            <%call userms(rs("user"),"mess")%> </td>
          <td width="616" valign="top" bgcolor="#FFFFFF"> <table width="100%" cellpadding="0" cellspacing="0">
              <tr bgcolor="#FFFFFF"> 
                <td width="91%" height="28" bgcolor="#E6E6E6"> <a href="#"> <img src="PIC/message.gif" width="60" height="18" border="0" onClick="MM_openBrWindow('sendmess.asp?user=<%=rs("user")%>','a','resizable=yes,width=500,height=450')"></a> 
                  <a href="usermanager.asp?p=addfriend&fd=<%=rs("user")%>"><img src="PIC/friend.gif" width="48" height="18" border="0"></a> 
                  <a href="usertype.asp?user=<%=rs("user")%>"><img src="PIC/profile.gif" width="45" height="18" border="0"></a> 
                  <a href="mailto:<%call userms(rs("user"),"email")%>"><img src="PIC/email.gif" width="45" height="18" border="0"></a> 
                  <a href="AddIn.asp?id=<%=id%>&lb=<%=lb%>&czlb=vote"><img src="PIC/reply.gif" width="45" height="18" border="0"></a> 
                  <img src="PIC/find.gif" width="45" height="18"> <a href="AddIn.asp?id=<%=id%>&lb=<%=lb%>&czlb=retz"><img src="PIC/reply_a.gif" width="45" height="18" border="0"></a></td>
                <td width="9%" align="center" bgcolor="#E6E6E6"><strong>��¥</strong></td>
              </tr>
              <tr> 
                <td colspan="2" bgcolor="#FFFFFF"><br> <TABLE cellSpacing=0 cellPadding=0 border=0>
                    <TBODY>
                      <TR> 
                        <TD width=14><IMG height=8 src="PIC/top_l.gif" 
            width=14></TD>
                        <TD background=PIC/top_c.gif></TD>
                        <TD width=16><IMG height=8 src="PIC/top_r.gif" 
            width=16></TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top width=14 background=PIC/center_l.gif></TD>
                        <TD style="FONT-SIZE: 9pt; LINE-HEIGHT: 12pt" bgColor=#f1f2f4> 
                          <%
response.Write("<b><img src="&rs("heat")&" alt=��������> "&name&"</b>")
if rs("lock")="locked" then
response.write("  <img src=""pic/locked.gif"" alt=""�������ѱ�����!�����ܻظ�������!""> (<font color=#FF0000>���������ܻظ�</font>)")
end if
response.Write("<br>")
if rs("kill")=1 then
response.write"<br>&nbsp;[ <font color=#FF0000>����������Ա���!</font> ]"
else
response.write ubbcode(rs("ly"))
end if
%> </TD>
                        <TD vAlign=top width=16 background=PIC/center_r.gif>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top width=14><IMG height=42 
              src="PIC/foot_l1.gif" width=14></TD>
                        <TD background=PIC/foot_c.gif><IMG height=42 
              src="PIC/foot_l3.gif" width=36></TD>
                        <TD align=right width=16><IMG height=42 
              src="PIC/foot_r.gif" width=16></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <br> <table width="98%" border="0">
                    <tr> 
                      <td width="14%" align="right" valign="top">&nbsp;</td>
                      <td width="86%"> <%call userms(rs("user"),"uqm")%> </td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="24" align="center" valign="middle" bgcolor="#FFFFFF">
		  <% if master then 
		  response.Write("<a href=look-ip.asp?id="&id&"&czlb=lookip><img src=PIC/ip.gif width=13 height=15 border=0 alt='�鿴�û�IP��Դ������<br>IP: "&rs("ip")&"'></a>")
		  else
		  response.Write("<a href=look-ip.asp?id="&id&"&czlb=lookip><img src=PIC/ip.gif width=13 height=15 border=0 alt='�鿴�û�IP��Դ������<br>IP: *.*.*.*'></a>")
		  end if
		  response.Write("&nbsp;"&rs("time"))
		  %></td>
          <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="0">
              <tr>
                <td width="68%" height="17">&nbsp;</td>
                <td width="32%" align="right">
				<%
				if userislogin  and (loginuser=rs("user")) then
				response.Write("<a href=""AddIn.asp?id="&id&"&lb="&lb&"&czlb=edittz""><img src=PIC/edit.gif width=45 height=18 border=0></a>")
				end if
                 if master then %>
                  <a href="caozuo.asp?lb=<%=lb%>&id=<%=id%>&czlb=����"><img src="PIC/lockedadmin.gif" alt="��������" width="18" height="18" border="0"></a> 
                  <a href="caozuo.asp?lb=<%=lb%>&id=<%=id%>&czlb=�ö�"><img src="PIC/topadmin.gif" alt="�ö���ȥ��" width="18" height="18" border="0"></a> 
                  <a href="caozuo.asp?id=<%=id%>&lb=<%=lb%>&czlb=ɾ��"><img src="PIC/delete.gif" alt="ɾ������" width="18" height="18" border="0"></a> 
                  <a href="caozuo.asp?id=<%=id%>&lb=<%=lb%>&czlb=����"><img src="PIC/jing.gif" alt="��������" width="18" height="18" border="0"></a> 
                  <a href="caozuo.asp?id=<%=id%>&lb=<%=lb%>&czlb=����"><img src="PIC/copy.gif" alt="�ƶ�����" width="18" height="18" border="0"></a> 
                  <a href="caozuo.asp?id=<%=id%>&lb=<%=lb%>&czlb=����"><img src="PIC/fong.gif" alt="��ջ������" width="18" height="18" border="0"></a> 
                  <%end if%>
				  </td>
              	  </tr>
            </table> </td>
        </tr>
<%
sub listtreerely()
dim rs,i,bgcolor
i=0
set rs=conn.execute("select * from rely where rid='"&id&"' order by id")
if rs.eof then
exit sub
end if
response.Write("<tr><td colspan=2 height=28 background=backimg/bg1.gif>&nbsp;&nbsp;<font color=#FFFFFF><b>*&nbsp;����Ŀ¼�ṹ</b></font></td></tr>")
do while ((not rs.eof) and (not rs.bof))
if (i mod 2)=0 then
bgcolor="#FFFFFF"
else
bgcolor="#EEEEEE"
end if
if rs("kill")=1 then
response.write"<tr><td colspan=2 bgcolor="&bgcolor&" height=24>&nbsp;&nbsp;�ظ�:&nbsp;<img src="&rs("heat")&">&nbsp;[ <font color=#FF0000>����������Ա���!</font> ]  <i>--<a href=usertype.asp?user="&rs("user")&">"&rs("user")&"</a>,"&rs("time")&"</i></td></tr>"
else
response.write"<tr><td colspan=2 bgcolor="&bgcolor&" height=24>&nbsp;&nbsp;�ظ�:&nbsp;<img src="&rs("heat")&">&nbsp;<a href=type.asp?id="&id&"&rlyid="&rs("id")&">"&ChkBadWords(left(rs("ly"),40))&"</a> <i>("&Len(rs("ly"))&"��)-- <a href=usertype.asp?user="&rs("user")&">"&rs("user")&"</a>,"&rs("time")&"</i></td></tr>"
end if
i=i+1
rs.movenext
loop
end sub
%>		
<%
if rs("rely")>0 then 
call listrely(id)
end if
sub listrely(cid)
dim rs,sql,ps,rlyid
ps=1
rlyid=request.QueryString("rlyid")
set rs=server.createobject("adodb.recordset")
if rlyid<>"" then
sqltext="select * from rely where id="&rlyid&" order by id"
else
if bbsskin=1 then
exit sub
else
sqltext="select * from rely where rid='"&cid&"' order by id"
end if
end if
rs.open sqltext,conn,1,1

dim MaxPerPage,text,checkpage,URLparameter,page,CurrentPage,i
MaxPerPage=10
URLparameter="&id="&id
if request("page")<>0 then
page=request("page")
else 
page=1
end if
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
do while not rs.eof 
%>
        <tr> 
          <td align="center" valign="top" bgcolor="#FFFFFF"><br>
            <%call userms(rs("user"),"mess")%> </td>
          <td valign="top" bgcolor="#FFFFFF"> <table width="100%" cellpadding="0" cellspacing="0">
              <tr bgcolor="#FFFFFF"> 
                <td width="89%" height="31" valign="middle" bgcolor="#E6E6E6"> <a href="#"> 
                  <img src="PIC/message.gif" width="60" height="18" border="0" onClick="MM_openBrWindow('sendmess.asp?user=<%=rs("user")%>','a','resizable=yes,width=500,height=450')"></a> 
                  <a href="usermanager.asp?p=addfriend&fd=<%=rs("user")%>"><img src="PIC/friend.gif" width="48" height="18" border="0"></a> 
                  <a href="usertype.asp?user=<%=rs("user")%>"><img src="PIC/profile.gif" width="45" height="18" border="0"></a> 
                  <a href="mailto:<%call userms(rs("user"),"email")%>"><img src="PIC/email.gif" width="45" height="18" border="0"></a> 
                  <a href="AddIn.asp?id=<%=id%>&rid=<%=rs("id")%>&lb=<%=lb%>&czlb=revote"><img src="PIC/reply.gif" width="45" height="18" border="0"></a> 
                  <img src="PIC/find.gif" width="45" height="18"> <a href="AddIn.asp?id=<%=id%>&lb=<%=lb%>&czlb=retz"><img src="PIC/reply_a.gif" width="45" height="18" border="0"></a></td>
                <td width="11%" align="center" valign="middle" bgcolor="#E6E6E6">
<%response.Write("<b> "&page&"</b> ��<b> "&(MaxPerPage-ps+1)&"</b> ¥")%>
                </td>
              </tr>
              <tr> 
                <td colspan="2" bgcolor="#FFFFFF"><br> <TABLE cellSpacing=0 cellPadding=0 border=0>
                    <TBODY>
                      <TR> 
                        <TD><IMG height=8 src="PIC/top_l.gif" 
            width=14></TD>
                        <TD background=PIC/top_c.gif></TD>
                        <TD><IMG height=8 src="PIC/top_r.gif" 
            width=16></TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top background=PIC/center_l.gif></TD>
                        <TD style="FONT-SIZE: 9pt; LINE-HEIGHT: 12pt" bgColor=#f1f2f4> 
                          <%
response.Write("<img src="&rs("heat")&" alt=��������>")
if rs("kill")=1 then
response.write"<br>&nbsp;[ <font color=#FF0000>�˻ظ����ѱ�����Ա���!</font> ]"
else
response.write ubbcode(rs("ly"))
end if
%> </TD>
                        <TD vAlign=top background=PIC/center_r.gif>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top><IMG height=42 
              src="PIC/foot_l1.gif" width=14></TD>
                        <TD background=PIC/foot_c.gif><IMG height=42 
              src="PIC/foot_l3.gif" width=36></TD>
                        <TD align=right><IMG height=42 
              src="PIC/foot_r.gif" width=16></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <table width="98%" border="0">
                    <tr> 
                      <td width="14%" align="right" valign="top">&nbsp;</td>
                      <td width="86%"> <%call userms(rs("user"),"uqm")%> </td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="24" align="center" valign="middle" bgcolor="#FFFFFF"> 
		  <% if master then 
		  response.Write("<a href=look-ip.asp?id="&id&"&rid="&rs("id")&"&czlb=lookip><img src=PIC/ip.gif width=13 height=15 border=0 alt='�鿴�û�IP��Դ������<br>IP: "&rs("ip")&"'></a>")
		  else
		  response.Write("<a href=look-ip.asp?id="&id&"&rid="&rs("id")&"&czlb=lookip><img src=PIC/ip.gif width=13 height=15 border=0 alt='�鿴�û�IP��Դ������<br>IP: *.*.*.*'></a>")
		  end if
		  response.Write("&nbsp;"&rs("time"))
		  %></td>
          <td bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0">
              <tr> 
                <td width="79%" height="17">&nbsp;</td>
			
                <td width="21%" align="right">
				
				 <% 
				if userislogin and (loginuser=rs("user")) then
				response.Write("<a href=""AddIn.asp?rid="&rs("id")&"&id="&id&"&lb="&lb&"&czlb=editretz""><img src=PIC/edit.gif width=45 height=18 border=0></a>")
				end if
				 if master then 
				 %>
                  <a href="caozuo.asp?id=<%=id%>&lb=<%=lb%>&rid=<%=rs("id")%>&czlb=ɾ��"> 
                  <img src="PIC/delete.gif" width="18" height="18" border="0"></a> 
                  <a href="caozuo.asp?id=<%=id%>&lb=<%=lb%>&rid=<%=rs("id")%>&czlb=���"> 
                  <img src="PIC/fong.gif" alt="��ջ������" width="18" height="18" border="0"></a> 
                  <%end if%></td>
              </tr>
            </table> </td>
        </tr>
        <%
rs.movenext
if ps=>MaxPerPage then exit do
ps=ps+1
loop
'��׼��ҳ�ӳ���URLparameterΪURL������
response.Write("<tr><td colspan=""2"" bgcolor=""#FFFFFF"">")
dim pages,soonhost,pagename,nx,x,ISselect,tenpages
pages=cint(request.QueryString("page"))
if pages=0 then
pages=1
end if
if pages>rs.pagecount then
pages=rs.pagecount
end if
pagename=request.ServerVariables("SCRIPT_NAME")

response.Write("<table width=100% border=0 align=center cellpadding=2 cellspacing=0 bgcolor=#92b9fb>")
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
response.Write("</select> </td></tr></table></td></tr>")
end sub
if bbsskin=1 then
call listtreerely()
end if
%></table>
      <%call tabletop(true,"down")%>
      <div align="right"><%call listchangeLt()%></div>
    </td>
<%
if userislogin then
%>
  <tr> 
    <td>
	<FORM name=FORM1 onsubmit="return form1_onsubmit()" action="AddIn.asp?lb=<%=lb%>&czlb=ok" method=post>
        <%call tabletop(true,"top")%>
        <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
          <tr> 
            <td height="27" colspan="2" background="backimg/bg1.gif" bgcolor="#92b9fb"> 
              <div align="center"><b><font color="#FFFFFF">���ٻظ�</font><font color="#000000" size="3"> 
                <input name="czlb" type="hidden" id="czlb" value="retz">
                <input name="id" type="hidden" id="id" value="<%=id%>">
                <input name="name" type="hidden" id="name" value="[�ظ�]:<%=name%>">
                <input name="lb" type="hidden" id="lb" value="<%=lb%>">
                </font></b></div></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="18%" height="20"> <div align="center">��������:</div></td>
            <td width="82%" rowspan="2"> <!--#include file="getubb.asp" -->
              <textarea cols="75" name="ly" rows="6" wrap="file" ></textarea>
              </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="18%"> <div align="center">(֧��<a href="helpubb.asp">UBB</a>����)<br>
                ���� 
                <script language=javascript >
<!--

function form1_onsubmit() 
{
   if(document.FORM1.ly.value.length<1)
 {
   alert("������ظ�������");
   document.FORM1.ly.focus();
   return false;
  }

}
ie = (document.all)? true:false
if (ie){
function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.sublit.submit();}}
}
//-->
</script>
              </div></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"> <div align="center"> <font color="#0000FF"> 
                <input name=sublit type=submit id="sublit"  value="�� ��">
                &nbsp;&nbsp; 
                <input type="reset" name="Reset" value="�� ��">
                </font></div></td>
          </tr>
        </table>
        <%call tabletop(true,"down")%>
      </form>
	
	  </td>
  </tr>
<%end if%>
  <tr> 
    <td></td>
  </tr>
  <tr> 
    <td>
<%
sub userms(user,disw)
dim rs,userjf,jb,jibie
jb=array("����","����","����","����","�Ķ�","���","����","�߶�","�˶�","�Ŷ�","ʮ��")
set rs=conn.execute("select ubbqm,myface,width,height,user,userjf,time,ftz,jibie,setuser1,email from userinfo where user='"&user&"'")
if not rs.eof then
select case(disw)
case "mess"
userjf=rs("userjf")
if userjf<=500 then
jibie=jb(0)
elseif userjf>500 and userjf=<1000 then
jibie=jb(1)
elseif userjf>1000 and userjf=<2000 then
jibie=jb(2)
elseif userjf>2000 and userjf=<3000 then
jibie=jb(3)
elseif userjf>3000 and userjf=<4000 then
jibie=jb(4)
elseif userjf>4000 and userjf=<5000 then
jibie=jb(5)
elseif userjf>5000 and userjf=<7500 then
jibie=jb(6)
elseif userjf>7500 and userjf=<10000 then
jibie=jb(7)
elseif userjf>10000 and userjf=<15000 then
jibie=jb(8)
elseif userjf>15000 and userjf=<20000 then
jibie=jb(9)
elseif userjf>20000 then
jibie=jb(10)
end if
response.write"<img src=../"&rs("myface")&" width="&rs("width")&" height="&rs("height")&"><br><br>"
response.write"<b>"&rs("jibie")&"</b><br>"
response.write "<table align=""center"" width=100 style=""filter:glow(color=red, strength=2)""><div align=""center"">"&rs("user")&"</div></table>"
response.write"�ȼ���"&jibie&"<br>"
response.write"���£�"&rs("ftz")&"<br>"
response.write"���֣�"&rs("userjf")&"<br>"
response.write"ע��ʱ�䣺<br>"&rs("time")
case "uqm"
response.write rs("ubbqm")
case "email"
   if rs("setuser1")=0 then
   response.write "����"
   else
   response.write rs("email")
   end if
end select
end if
end sub

sub chanagego(gid,glb)
dim rs
set rs=conn.execute("SELECT id FROM borecorder where lb='"&glb&"' and id<"&gid&" order by id desc")
if not rs.eof then
response.Write("<a href=type.asp?id="&rs("id")&"><img src=pic/prethread.gif width=52 height=12 border=0></a></td><td width=""6%"">")
else
response.Write("<img src=pic/prethread.gif width=52 height=12 border=0 alt=�Ѿ�û����һ����></td><td width=""6%"">")
end if
response.Write("<a href=javascript:this.location.reload()><img src=pic/refresh.gif alt=ˢ�²�Ʒ width=40 height=12 border=0></a></td><td width=""8%"">")

set rs=conn.execute("SELECT id FROM borecorder where lb='"&glb&"' and id>"&gid&" order by id")
if not rs.eof then
response.Write("<a href=type.asp?id="&rs("id")&"><img src=pic/nextthread.gif width=52 height=12 border=0></a>")
else
response.Write("<img src=pic/prethread.gif width=52 height=12 border=0 alt=�Ѿ�û����һ����>")
end if

if bbsskin=1 then
response.Write("&nbsp;<a href=type.asp?id="&id&"&skin=0><img src=pic/flatview.gif border=0 alt=ƽ��ṹ��ʾ></a>")
else
response.Write("&nbsp;<a href=type.asp?id="&id&"&skin=1><img src=pic/treeview.gif border=0 alt=���ͽṹ��ʾ></a>")
end if
response.Write("</td></tr>")
end sub
%> </td>
  </tr>
</table>
<%
sub typeusertp(tpid)
dim rs
set rs=conn.execute("select * from lttp where id="&cint(tpid))
%>
<table width="770" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
  <form name="formtp" method="post" action="type.asp?id=<%=id%>">
<%
dim alltp,ltsel,tpcount,n,i,user,useriftp
dim tpuser
useriftp=false
'on error resume next
alltp=0
		ltsel=split(rs("tply"),"|")
        tpcount=split(rs("count"),"|")
if (rs("tpuser")<>"") and userislogin then
tpuser=split(rs("tpuser"),"|")
for n=0 to ubound(tpuser)
 if loginuser=tpuser(n) then
 useriftp=true
 end if
next
end if

for n=0 to ubound(tpcount)
if tpcount(n)<>"" then
alltp=alltp+Cint(tpcount(n))
end if
next
		for i=0 to ubound(ltsel)
if ltsel(i)<>"" then%>
    <tr bgcolor="#FFFFFF"> 
      <td width="251">
      <%=i+1%><input name=tp type=radio value=<%=i%>><%=dvHTMLEncode(ltsel(i))%></td> 
      <td>
	  <%dim bfb
	  if alltp=0 then
	  bfb=0
	  else
	  bfb=(tpcount(i)/alltp)
	  end if%>
	  <img src="../obbs/PIC/bar2.gif" width="<%=bfb*300%>" height="10"> 
        <%=left((bfb*100),5)%>% &nbsp;&nbsp;<%=tpcount(i)%> Ʊ</td>
    </tr>
    <%
end if 
next
%>
    <tr bgcolor="#FFFFFF"> 
      <td height="18" colspan="2">
	  <%
	  if userislogin and (not useriftp) then 
	  response.write("<input type=submit name=Submit value=ͶƱ><input name=id2 type=hidden value="&rs("id")&">")
      else
	  response.Write("���Ѿ�Ͷ��Ʊ�˻���δ��¼,���Բ��ܲμ�ͶƱ!")
	  end if
	  %>
        <input name="tpid" type="hidden" id="tpid" value="<%=tpid%>"> </td>
    </tr>
  </form>
</table>
<%
end sub
sub adduserlttp()
if not userislogin then
founderr=true
errmess=errmess&error1
end if
dim sqltext,rs,ttuser,i

set rs=server.createobject("adodb.recordset")
sqltext="select tpuser,count,click,userselect from lttp where id="&request.Form("tpid")
rs.open sqltext,conn,1,3
if rs("tpuser")<>"" then
ttuser=split(rs("tpuser"),"|")
 for i=0 to ubound(ttuser)
 if ttuser(i)<>"" and ttuser(i)=loginuser then
 response.write "<script language=JavaScript>" & chr(13) & "alert('���Ѿ�Ͷ��Ʊ��.�㲻����ͶƱ��.�Ǻ�!');history.back()</script>"
 exit sub
 end if
 next
end if

dim tp,tpcount,tpz
tp=cint(request.Form("tp"))
tpcount=split(rs("count"),"|")
  for i=0 to ubound(tpcount)
if tpcount(i)<>"" then
  if i=tp then
  tpcount(i)=tpcount(i)+1
  end if
  tpcount(i)=tpcount(i)&"|"  
  tpz=tpz&tpcount(i)
end if
  next
rs("count")=tpz
rs("click")=rs("click")+1
rs("userselect")=request.Form("tp")&"|"&rs("userselect")
rs("tpuser")=loginuser&"|"&rs("tpuser")
rs.update
end sub
%>
<!--#include file="end.asp" -->
</body>
</html>
