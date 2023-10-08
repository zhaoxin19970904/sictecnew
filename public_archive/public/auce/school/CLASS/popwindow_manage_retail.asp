<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
curclassid=trim(session("classid"))
if userid="" then
    response.redirect "../error.asp?info=对不起，您已经掉线了，请重新申请！"
end if
set rs = createobject("ADODB.recordset")
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
   monitor=rs("monitor")
   monitor1=rs("monitor1")
end if
rs.close
if userid<>monitor then
   response.redirect "../error.asp?info=对不起，您不是本班正管理员无权进入！"
end if  
%> 
<HTML>
<HEAD>
<title>班级管理--辞职</title>
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
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" class="topic"><font color="#CC0000">班级管理</font></td>
        </tr>
        <tr> 
          <td bgcolor="#B4C7D4" height="1"></td>
        </tr>
      </table>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><a href="popwindow_manage_deletecy.asp">驱逐成员</a> | <a href="popwindow_manage_emailall.asp">给全体成员发信</a> 
            | <a href="popwindow_manage_noticeall.asp">发布班级通告</a> | <a href="popwindow_manage_takesec.asp">任命副管理员</a> 
            | <a href="popwindow_manage_retail.asp">辞职</a></td>
        </tr>
      </table>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>辞职---</td>
        </tr>
      </table>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <form name="form1" method="post" action="retail.asp"><tr>
            <td height="20" valign="top" class="topic2">感谢你对班级管理做出的贡献，请填写好新管理员的用户名后，点击辞职！</td>
        </tr>
        <tr>
          <td align="center">新管理员用户名：
            
              <input type="text" name="monitor" size="20">
              <input type="submit" name="Submit" value="辞职!">
            </td>
        </tr> </form>
      </table>
   
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>