<!--#include file="conn.asp" -->
<!--#include file="ubbcode.asp" -->
<!--#include file="mymem.asp" -->
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
%>
<table width="898" border=0 align="center" cellpadding=0 cellspacing=0 bgcolor="#FFFFFF">
  <tbody>
    <tr> 
      <td height="10" bgcolor="#FFFFFF" >&nbsp;</td>
    </tr>
    <tr> 
      <td height="21"> 
        <%
	 call ltmessetruyg()
	 sub ltmessetruyg()
	 dim rs,retzsums,tzsums,newuser,todaytz,yestdayft
	 tzsums=0
	 retzsums=0
	 set rs=conn.execute("select top 1 user from userinfo order by id desc")
	 newuser=rs("user")
	 set rs=conn.execute("select tzsum,retzsum,todayft,yestdayft from ltlb")
	 do while not rs.eof 
	 retzsums=retzsums+rs("retzsum")
	 tzsums=tzsums+rs("tzsum")
	 todaytz=todaytz+rs("todayft")
	 yestdayft=yestdayft+rs("yestdayft")
	 rs.movenext
	 loop
	 response.Write("��ӭ�¼����Ա <b><a href=usertype.asp?user="&newuser&">"&newuser&"</a></b><br><br>")
	 response.Write("��̳����<strong>: "&loadcopyc("reguser")&" </strong><a href=listalluser.asp>ע���Ա</a>,")
	 response.Write("��������<b>: "&tzsums&"</b>,��������<b>: "&(retzsums+tzsums))&"</b><br><br>"
	 response.Write("���չ�������<b>: <font color=red>"&todaytz&"</font></b>,��������<b>: "&yestdayft&"</b>,����շ���: ")
	 set rs=conn.execute("select hft,hfttime from admin_copyc")
	 if rs("hft")<todaytz then
	 conn.execute("update admin_copyc set [hft]="&todaytz&",[hfttime]=now()")
	 end if
	 response.write("<b>"&rs("hft")&" </b> ["&rs("hfttime")&"]")
	 end sub
	 %>
      </td>
    </tr>
    <tr> 
      <td height="18"> 
        <%
call listzlt()
sub listzlt()
dim rs,sql,i
 if request.QueryString("atlb")<>"" then '��ʾ��������̳
 set rs=conn.execute("select id,ltname from ltname where id="&request.QueryString("atlb"))
 call listlbboard(rs("id"),rs("ltname"))
 else '��ʾ���з���̳
 Set rs = Server.CreateObject("ADODB.Recordset")
 sql="select id,ltname from ltname order by b_order desc,id desc" 
 rs.open sql,conn,1,1
 for i=1 to rs.recordcount
 call listlbboard(rs("id"),rs("ltname"))
 rs.movenext
 next
 end if
rs.close
set rs=nothing
end sub
sub listlbboard(zlt,ltname) '��ʾ����̳�ӳ���
dim rs,sql,i,ltconfig
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from ltlb where atlb='"&zlt&"'order by b_order desc,id desc" 
rs.open sql,conn,1,1
call tabletop(true,"top")
%>
        <table width="100%" align="center" bgcolor="#92b9fb" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="#0053A6"> 
            <td height="27" background="backimg/bg1.gif" bgcolor="#0053A6" > <img id="isimg<%=zlt%>" onClick="HidexMenu('<%=zlt%>')" src="backimg/nofollow.gif" width="15" height="15"><b><font color="#FFFFFF"><a href=index.asp?atlb=<%=zlt%>><%=ltname%></a> <a href="javascript:HidexMenu('<%=zlt%>')">[��ȫ�ر�]</a></font></b></td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#EEEEEE" id="tr<%=zlt%>"> 
              <table  width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#92b9fb">
                <%for i=1 to rs.recordcount%>
                <tr> 
                  <td width="10%" align="center" bgcolor="#EEEEEE">
				  <%
				  ltconfig=split(rs("ltconfig"),"|")
				  if ltconfig(0)="1" then
				  response.Write("<img src=backimg/member.gif>")
				  elseif ltconfig(1)="1" then
				  response.Write("<img src=backimg/lock.gif>")
				  elseif rs("todayft")>0 then
				  response.Write("<img src=backimg/rz.gif>")
				  else
				  response.Write("<img src=backimg/on.gif>")
				  end if
				  %>
                  </td>
                  <td bgcolor="#FFFFFF"> <table width="100%" border="0" cellpadding="3" cellspacing="0">
                      <tr> 
                        <td width="65%" valign="top"> 
                          <%
					response.Write("<a href=list.asp?lb="&rs("id")&">"&rs("ltlb")&"</a>")
					if rs("ifoff")=1 then
					response.Write(" -- <font color=red>��ʱ�ر�</font>")
					end if
					response.Write("<p>&nbsp;&nbsp;<img src=backimg/Forum_readme.gif> "&rs("ltsm")&"</p>")
					%>
                        </td>
                        <td width="26%"> 
                          <%call lastboardly(rs("id"))%>
                          <img src="backimg/lastpost.gif" width="11" height="10"> 
                        </td>
                      </tr>
                      <tr> 
                        <td height="32" bgcolor="#EEEEEE">����: 
<%
dim ltsel,x
if rs("ltbz")<>"" then
ltsel=split(rs("ltbz"),"|")
for x=0 to (ubound(ltsel)-1)
'if  userqx(ltsel(x))="����" then
response.write" * <a href=usertype.asp?user="&ltsel(x)&" title=�鿴������Ϣ>" &ltsel(x)&"</a>"
'end if
next
else
response.Write("���ް���")
end if
%>
                        </td>
                        <td align="right" bgcolor="#EEEEEE"><img src="PIC/Forum_today.gif" alt="��������" width="25" height="10"> 
                          <font color="#FF0000"> <%=rs("todayft")%> </font> <img src="PIC/forum_topic.gif" alt="����" width="25" height="10"> 
                          <%=rs("tzsum")%> <img src="PIC/Forum_post.gif" alt="�ظ���" width="25" height="10"> 
                          <%=rs("retzsum")%>&nbsp;&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <%
		  rs.movenext 
		  next
		  %>
              </table></td>
          </tr>
        </table>
<%
call tabletop(true,"down")
response.Write("<br>")
end sub
response.Write("<div align=right>")
call listchangeLt()
response.Write("</div>")
%> 
</td>
    </tr>
    <tr> 
      <td height="17"><br></td></tr>
    <tr>
      <td height="17"><table width="100%" align="center" bgcolor="#92b9fb" border="0" cellpadding="3" cellspacing="1">
          <tr bgcolor="#0053A6"> 
            <td height="27" background="backimg/bg1.gif" bgcolor="#0053A6" > <b><font color="#FFFFFF">-=&gt; 
              �����б� <a href=loadtree.asp?listonline=user target="hiddenframex" onClick="document.all.listonline.innerHTML='���ڶ�ȡ�û��б����Ժ��'">[��ʾ��ϸ�б�]</a></font></b>
			 <IFRAME name="hiddenframex"  frameBorder=0 width=0 height=0></IFRAME>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" > <%call onlineusersum("all")%>
              <br>
              <font color="#FF0000">����ͼ��</font> ����Ա <img src="PIC/ao1.gif" width="16" height="18"> 
              �ο� <img src="PIC/messages1.gif" width="12" height="11"> 
              <hr align="left" width="100%" size="1" noshade color="#92b9fb">
             <div id="listonline"></div> </td>
          </tr>
        </table></td>
    </tr><tr>
	  <td align="center"><br>
        <img src="backimg/on.gif"> û�������� <img src="backimg/lock.gif" width="28" height="27"> 
        ������̳<img src="backimg/member.gif"> ��Ա��̳ <img src="backimg/rz.gif" width="28" height="27"> 
        ���µ�����</td>
    </tr>
</table>

<SCRIPT language=JavaScript>
function listonlinediv(){
 document.frames["hiddenframe"].location.replace("loadtree.asp?listonline=user");
// if(document.all["listonline"].style.display=="none"){
//   		 document["listonline"].style.display=""}
//	else{
//         document.all["listonline"].style.display="none"
//		}
}

function HidexMenu(thenum){
var thetool
var theimg
thetool="tr"+thenum
theimg="isimg"+thenum
 	if (document.all[thetool].style.display=="none"){ 
    document.all[thetool].style.display="";
	document.images[theimg].src="backimg/nofollow.gif"}
 	else{
    document.all[thetool].style.display ="none";
	document.images[theimg].src="backimg/follow.gif"} 		
} 
</SCRIPT>
<% 
sub lastboardly(zlt)
dim rs,sql,i,rtzs
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select user,id,time,rely,name from borecorder where lb='"&zlt&"'order by id desc"
rs.open sql,conn,1,1
if not rs.eof and not rs.bof then
response.Write("����:<a href=type.asp?id="&rs("id")&">"&dvHTMLEncode(left(rs("name"),11))&"</a>...<br>����:"&rs("user")&"<br>ʱ��:"&rs("time"))
end if
end sub
%> 
<!--#include file="end.asp" -->

</body>
</html>
