<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
curclassid=trim(session("classid"))
if userid="" then
    response.redirect "../error.asp?info=�Բ������Ѿ������ˣ����������룡"
end if
set rs = createobject("ADODB.recordset")
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
   monitor=rs("monitor")
   monitor1=rs("monitor1")
end if
rs.close
if userid<>monitor or userid<>monitor1 then
   response.redirect "../error.asp?info=�Բ��������Ǳ������Ա��Ȩ���룡"
end if  

%> 
<HTML>
<HEAD>
<title>�༶����--�����༶ͨ��</title>
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
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" class="topic"><font color="#CC0000">�༶����</font></td>
        </tr>
        <tr> 
          <td bgcolor="#B4C7D4" height="1"></td>
        </tr>
      </table>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><a href="popwindow_manage_deletecy.asp">�����Ա</a> | <a href="popwindow_manage_emailall.asp">��ȫ���Ա����</a> 
            | <a href="popwindow_manage_noticeall.asp">�����༶ͨ��</a> | <a href="popwindow_manage_takesec.asp">����������Ա</a> 
            | <a href="popwindow_manage_retail.asp">��ְ</a></td>
        </tr>
      </table>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>�����༶ͨ��---</td>
        </tr>
      </table>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <form name="form1" method="post" action="tonggao.asp">
          <tr> 
            <td class="topic2" valign="middle">��&nbsp;&nbsp;�ݣ�</td>
            <td class="topic" valign="top"> 
              <textarea name="tonggao" cols="65" rows="10"></textarea>
            </td>
          </tr>
          <tr valign="bottom" align="center"> 
            <td class="di" height="35" colspan="2"> 
              <input type="submit" name="Submit" value="����">
              &nbsp;&nbsp; 
              <input type="reset" name="Submit3" value="��д">
            </td>
          </tr>
        </form>
      </table>
      <br>
      <br>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>