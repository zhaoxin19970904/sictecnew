 <!--#include file="conn.asp" -->
<!--#include file="ubbcode.asp" -->
<!--#include file="mymem.asp" -->
<table width="898" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td colspan="2"> 
      <%call tabletop(true,"top")%>
    </td>
  </tr>
  <tr> 
    <td height="27" colspan="2" background="backimg/bg1.gif"><strong><font color="#FFFFFF"> 
      <%=request.QueryString("user")%>&nbsp;&nbsp;�û���Ϣ </font></strong></td>
  </tr>
  <tr> 
    <td colspan="2" valign="top"> <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">

        <tr bgcolor="#FFFFFF"> 
          <td> <table width="100%" border="0">
              <tr> 
          <td height="98" valign="top">
		  <%
		  call listuserjbmess(request.QueryString("user"))
		  sub listuserjbmess(user)
		  if user="" then
		  response.cookies("rewin")("errmess")="<li>�Բ����û�!����Ϊ�� �����ȷ�����ӽ���!<a href=index.asp> ������ҳ</a>"
		  response.Redirect("error.asp?founderr=mess")
		  end if
		  dim rs,rs2
		  set rs=conn.execute("select * from userinfo where user='"&user&"'")
		  if rs.eof or rs.bof then
		  response.cookies("rewin")("errmess")="<li>���ٴ���!����û���ҵ����û�����! �����ȷ�����ӽ���!<a href=index.asp> ������ҳ</a>"
		  response.Redirect("error.asp?founderr=mess")
		  end if
		  %> 
                  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
                    <tr align="right" bgcolor="#FFFFFF"> 
                      <td colspan="3"><img src="PIC/zhuangtai.gif" width="16" height="16"> 
                        ״̬�� 
                        <%
if rs("kill")=0 then
response.Write(" ���� ")
else
response.Write(" ID����� ")
end if
set rs2=conn.execute("select user from online where user='"&user&"'")
  if not rs2.eof then
  response.Write(" [����]")
  else
  response.Write(" [����]")
  end if
%> </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">����</td>
                      <td width="55%"><%=rs("jibie")%></td>
                      <td width="30%" rowspan="8" align="center"><img src="<%=rs("myface")%>" width="<%=rs("width")%>" height="<%=rs("height")%>" id=idface><br> 
                        <br>
                        �û�ͷ��</td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">�û���</td>
                      <td><%=rs("user")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">�� ��</td>
                      <td><%=rs("xb")%>&nbsp;</td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">�� ��</td>
                      <td><%=rs("birthday")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">E-mail</td>
                      <td> <%if rs("setuser1")=0 then response.write "�û�δ����" else response.write rs("email") end if%> </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">oicq</td>
                      <td><%=rs("oicq")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">��������</td>
                      <td><%=rs("ftz")%>��</td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">��ɾ����</td>
                      <td><%=rs("deltz")%>��</td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">��½����</td>
                      <td><%=rs("userlog")%> ��</td>
                      <td align="center">�������� | <a href="usermanager.asp?p=addfriend&fd=<%=rs("user")%>">��Ϊ����</a> 
                        | <a href="look-ip.asp?user=<%=rs("user")%>&czlb=userip">IP��Դ</a></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">�����˳�</td>
                      <td colspan="2"><%=rs("logout")%> ��</td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">���ڻ���</td>
                      <td colspan="2"><%=rs("userjf")%> ��</td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">��ϵ��ʽ</td>
                      <td colspan="2"> <%if rs("setuser1")=0 then response.write "�û�δ����" else response.write rs("conn") end if%> </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">����ʱ��</td>
                      <td colspan="2"> <%=rs("atime")%> ��</td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td width="15%">ע��ʱ��:</td>
                      <td colspan="2"><%=rs("time")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%">���һ�ε�¼:</td>
                      <td colspan="2"><%=rs("lastlogin")%></td>
                    </tr>
                  </table>
                  <%end sub%> </td>
              </tr>
              <tr>
                <td height="98" valign="top"><br>
<%
call listzjftzx(10,request.QueryString("user"))
sub listzjftzx(totles,username)
dim rs,i
set rs=conn.execute("select id,name,time,user,retime,lastly,click,rely,lb from borecorder where user='"&username&"' order by id desc",300,1)
%> 

                  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
                    <tr height="27" background="backimg/bg1.gif"> 
                      <td width="43%" height="27" background="backimg/bg1.gif"> 
                        <strong>��������������</strong></td>
                      <td width="12%" align="center" background="backimg/bg1.gif">����</td>
                      <td width="12%" align="center" background="backimg/bg1.gif">����/�ظ�</td>
                      <td width="11%" align="center" background="backimg/bg1.gif">���ظ�</td>
                      <td width="22%" align="center" background="backimg/bg1.gif">����ʱ��</td>
                    </tr>
             <%
			 if rs.eof or rs.bof then 
			  response.Write("<tr><td bgcolor=#FFFFFF colspan=""5"">�㻹û�в�����κ�����</td></tr></table>") 
			 exit sub
			 end if
			  if totles=10 then 
			  do while not rs.eof
			  response.Write("<tr bgcolor=#FFFFFF><td>�� <a href=type.asp?id="&rs("id")&">"&dvHTMLEncode(rs("name"))&"</a></td><td align=center>"&rs("user")&"</td><td align=center>"&rs("click")&"/"&rs("rely")&"</td><td align=center>"&rs("lastly")&"</td><td align=center>"&rs("time")&"</td></tr>")
			  rs.movenext
			  i=i+1
			  if i>10 then exit do
			  loop
			  else
			  do while not rs.eof
			  response.Write("<tr bgcolor=#FFFFFF><td>�� <a href=type.asp?id="&rs("id")&">"&dvHTMLEncode(rs("name"))&"</a>")
        '���ظ��ﵽ��ҳʱ��ʾ
		dim crely,x,lb
		lb=rs("lb")
		crely=rs("rely")
		if crely>10 then
		dim disppagesd
		response.Write("[<b><img src=backimg/multipage.gif>")
		if CInt(crely/10+0.5)>8 then 
		disppagesd=8
		else
		disppagesd=CInt(crely/10+0.5)
		end if
        for x=1 to disppagesd
        Response.write(" <a href='type.asp?page="&x&"&lb="&lb&"&id="&rs("id")&"'><font color=red>"&x&"</font></a>")
        next
		if disppagesd=8 then 
		response.Write(" ... <a href='type.asp?page="&CInt(crely/10+0.5)&"&lb="&lb&"&id="&rs("id")&"'><font color=red>"&CInt(crely/10+0.5)&"</font></a>")
		end if
		response.Write("</b>]")
        end if			  
 			  response.Write("</td><td align=center>"&rs("user")&"</td><td align=center>"&rs("click")&"/"&rs("rely")&"</td><td align=center>"&rs("lastly")&"</td><td align=center>"&rs("time")&"</td></tr>")  
			  rs.movenext
			  i=i+1
			  if i>300 then exit do
			  loop
			 end if
%>
                  </table>
<%end sub%></td>
              </tr>
            </table></td>
        </tr>

      </table></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <%call tabletop(true,"down")%>
    </td>
  </tr>
</table>
<!--#include file="end.asp" -->
