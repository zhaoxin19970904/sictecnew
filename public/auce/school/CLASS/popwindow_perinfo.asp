<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
if userid="" then
    response.redirect "../error.asp?info=�Բ������Ѿ������ˣ����������룡"
end if
curuserid=trim(request("userid"))
if curuserid<>"" then
   userid=curuserid
end if
set rs = createobject("ADODB.recordset")  
if curuserid="" then  
   sql="select * from userinfo where userid='"&userid&"'"
   rs.open SQL,schooldb
   if not rs.eof then
      realname=rs("realname")
      dearname=rs("dearname")
      userbirth=rs("userbirth")
   end if
   rs.close  
   sql="select * from usercommunicationinfo where userid='"&userid&"'"
   rs.open SQL,schooldb
   if not rs.eof then
      communicationaddr=rs("communicationaddr")
      telephone=rs("telephone")
      mobile=rs("mobile")
      bp=rs("bp")
      qq=rs("qq")
      workshop=rs("workshop")
      homeaddr=rs("homeaddr")
      email=rs("email")
   end if
   rs.close 
%>
<HTML>
<HEAD>
<title>������Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0"><!--#include file=../top1.htm-->
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="top"> <br>
      <table width="90%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td class="topic" height="20" valign="top"><font color="#000000">������Ϣ</font></td>
        </tr>
      </table>
      <table width="500" border="0" cellspacing="1" cellpadding="1">
   
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">��ʵ������</font><br>
            </td>
            <td height="15" width="370" class="topic"><%=realname%></td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�ǳƣ���������</font></td>
            <td width="370" class="topic"><%=dearname%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�������գ�</font></td>
            <td width="370" class="topic"><%=userbirth%></td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">ͨ�ŵ�ַ(�ʱ�)��</font></td>
            <td width="370" class="topic"><%=communicationaddr%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�̶��绰��</font></td>
            <td width="370" class="topic"><%=telephone%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�ƶ��绰��</font></td>
            <td width="370" class="topic"><%=mobile%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">������</font></td>
            <td width="370" class="topic"><%=bp%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">OICQ/ICQ ���룺</font></td>
            <td width="370" class="topic"><%=qq%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">������λ��</font></td>
            <td width="370" class="topic"><%=workshop%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">��ס��ַ��</font></td>
            <td width="370" class="topic"><%=homeaddr%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�����ʼ���ַ��</font></td>
            <td width="370" class="topic"><%=email%> </td>
          </tr>
          <tr align="center" valign="bottom"> 
            <td colspan="2" class="topic" height="35"><a href="javascript:window.close();">
            <font color="#000000">�رմ���</font></a></td>
          </tr>

      </table>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>
<%
else
   sql="select * from userjoinclassinfo where userid='"&curuserid&"' and classid='"&trim(session("classid"))&"'"
   rs.open SQL,schooldb
   if rs.eof then
      response.redirect "../error.asp?info=�Բ�������Ȩ�鿴������Ϣ��"
   end if
   rs.close
    sql="select * from userinfo where userid='"&userid&"'"
   rs.open SQL,schooldb
   if not rs.eof then
      realname=rs("realname")
      dearname=rs("dearname")
      userbirth=rs("userbirth")
   end if
   rs.close  
   sql="select * from usercommunicationinfo where userid='"&userid&"'"
   rs.open SQL,schooldb
   if not rs.eof then
      communicationaddr=rs("communicationaddr")
      telephone=rs("telephone")
      mobile=rs("mobile")
      bp=rs("bp")
      qq=rs("qq")
      workshop=rs("workshop")
      homeaddr=rs("homeaddr")
      email=rs("email")
   
    
%>
<HTML>
<HEAD>
<title>������Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0"><!--#include file=../top1.htm-->
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="top"> <br>
      <table width="90%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td class="topic" height="20" valign="top"><font color="#000000">������Ϣ</font></td>
        </tr>
      </table>
      <table width="500" border="0" cellspacing="1" cellpadding="1">
   
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">��ʵ������</font><br>
            </td>
            <td height="15" width="370" class="topic"><%=realname%></td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�ǳƣ���������</font></td>
            <td width="370" class="topic"><%=dearname%><font color="#000000">
            </font> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�������գ�</font></td>
            <td width="370" class="topic"><%=userbirth%></td>
          </tr>
          <%if rs("iscashow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">ͨ�ŵ�ַ(�ʱ�)��</font></td>
            <td width="370" class="topic"><%=communicationaddr%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">ͨ�ŵ�ַ(�ʱ�)��</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%> 
          <%if rs("istelephoneshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�̶��绰��</font></td>
            <td width="370" class="topic"><%=telephone%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�̶��绰��</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%> 
           <%if rs("ismobileshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�ƶ��绰��</font></td>
            <td width="370" class="topic"><%=mobile%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�ƶ��绰��</font></td>
            <td width="370" class="topic"><font color="#000000">**********</font></td>
          </tr>
          <%end if%>
           <%if rs("isbpshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">������</font></td>
            <td width="370" class="topic"><%=bp%><font color="#000000"> </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">������</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%>
          <%if rs("isqqshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">OICQ/ICQ ���룺</font></td>
            <td width="370" class="topic"><%=qq%><font color="#000000"> </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">OICQ/ICQ ���룺</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%>
          <%if rs("iswsshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">������λ��</font></td>
            <td width="370" class="topic"><%=workshop%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">������λ��</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%>
           <%if rs("ishashow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">��ס��ַ��</font></td>
            <td width="370" class="topic"><%=homeaddr%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">��ס��ַ��</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%>
          <%if rs("isemailshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�����ʼ���ַ��</font></td>
            <td width="370" class="topic"><%=email%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">�����ʼ���ַ��</font></td>
            <td width="370" class="topic"><font color="#000000">**********</font></td>
          </tr>
          <%end if%>
          <tr align="center" valign="bottom"> 
            <td colspan="2" class="topic" height="35"><a href="javascript:window.close();">
            <font color="#000000">�رմ���</font></a></td>
          </tr>

      </table>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>   
<%
end if 
rs.close
end if
%>