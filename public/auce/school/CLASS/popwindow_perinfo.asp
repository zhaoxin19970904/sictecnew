<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
if userid="" then
    response.redirect "../error.asp?info=对不起，您已经掉线了，请重新申请！"
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
<title>个人信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
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
          <td class="topic" height="20" valign="top"><font color="#000000">个人信息</font></td>
        </tr>
      </table>
      <table width="500" border="0" cellspacing="1" cellpadding="1">
   
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">真实姓名：</font><br>
            </td>
            <td height="15" width="370" class="topic"><%=realname%></td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">昵称（笔名）：</font></td>
            <td width="370" class="topic"><%=dearname%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">您的生日：</font></td>
            <td width="370" class="topic"><%=userbirth%></td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">通信地址(邮编)：</font></td>
            <td width="370" class="topic"><%=communicationaddr%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">固定电话：</font></td>
            <td width="370" class="topic"><%=telephone%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">移动电话：</font></td>
            <td width="370" class="topic"><%=mobile%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">传呼：</font></td>
            <td width="370" class="topic"><%=bp%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">OICQ/ICQ 号码：</font></td>
            <td width="370" class="topic"><%=qq%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">工作单位：</font></td>
            <td width="370" class="topic"><%=workshop%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">居住地址：</font></td>
            <td width="370" class="topic"><%=homeaddr%> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">电子邮件地址：</font></td>
            <td width="370" class="topic"><%=email%> </td>
          </tr>
          <tr align="center" valign="bottom"> 
            <td colspan="2" class="topic" height="35"><a href="javascript:window.close();">
            <font color="#000000">关闭窗口</font></a></td>
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
      response.redirect "../error.asp?info=对不起，您无权查看他的信息！"
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
<title>个人信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
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
          <td class="topic" height="20" valign="top"><font color="#000000">个人信息</font></td>
        </tr>
      </table>
      <table width="500" border="0" cellspacing="1" cellpadding="1">
   
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">真实姓名：</font><br>
            </td>
            <td height="15" width="370" class="topic"><%=realname%></td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">昵称（笔名）：</font></td>
            <td width="370" class="topic"><%=dearname%><font color="#000000">
            </font> </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">您的生日：</font></td>
            <td width="370" class="topic"><%=userbirth%></td>
          </tr>
          <%if rs("iscashow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">通信地址(邮编)：</font></td>
            <td width="370" class="topic"><%=communicationaddr%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">通信地址(邮编)：</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%> 
          <%if rs("istelephoneshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">固定电话：</font></td>
            <td width="370" class="topic"><%=telephone%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">固定电话：</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%> 
           <%if rs("ismobileshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">移动电话：</font></td>
            <td width="370" class="topic"><%=mobile%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">移动电话：</font></td>
            <td width="370" class="topic"><font color="#000000">**********</font></td>
          </tr>
          <%end if%>
           <%if rs("isbpshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">传呼：</font></td>
            <td width="370" class="topic"><%=bp%><font color="#000000"> </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">传呼：</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%>
          <%if rs("isqqshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">OICQ/ICQ 号码：</font></td>
            <td width="370" class="topic"><%=qq%><font color="#000000"> </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">OICQ/ICQ 号码：</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%>
          <%if rs("iswsshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">工作单位：</font></td>
            <td width="370" class="topic"><%=workshop%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">工作单位：</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%>
           <%if rs("ishashow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">居住地址：</font></td>
            <td width="370" class="topic"><%=homeaddr%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">居住地址：</font></td>
            <td width="370" class="topic"><font color="#000000">**********
            </font> </td>
          </tr>
          <%end if%>
          <%if rs("isemailshow")=1 then%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">电子邮件地址：</font></td>
            <td width="370" class="topic"><%=email%><font color="#000000">
            </font> </td>
          </tr>
          <%else%>
          <tr> 
            <td width="130" align="right" class="topic" height="20">
            <font color="#000000">电子邮件地址：</font></td>
            <td width="370" class="topic"><font color="#000000">**********</font></td>
          </tr>
          <%end if%>
          <tr align="center" valign="bottom"> 
            <td colspan="2" class="topic" height="35"><a href="javascript:window.close();">
            <font color="#000000">关闭窗口</font></a></td>
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