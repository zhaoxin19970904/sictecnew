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
<title>�༶����--�����Ա</title>
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
          <td>�����Ա---</td>
        </tr>
      </table>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" class="topic" width="300"><font color="#CC0000">��Ա</font></td>
          <td height="20" class="digi" width="210" align="left" valign="middle">��</td>
        </tr>
        <tr> 
          <td bgcolor="#B4C7D4" height="1" colspan="2"></td>
        </tr>
      </table>
      <%
            if Request("Page")<>"" then 
       		Page = CLng(Request("Page"))
    	    end if 
            If Page < 1 Then 
                Page = 1
            end if
       %>     
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <%
                    PageSize=10
                    set rs=Server.CreateObject("ADODB.RecordSet")
                    sql="select count(*) as a from userjoinclassinfo where classid ='"&curclassid&"'"  
                    rs.open SQL,schooldb,1,3
                    count=rs("a")
                    PageCount=CInt(rs("a")/PageSize+0.5)
	            rs.Close
	            if page=1 then
	       SQL="select top 10 * from userjoinclassinfo where classid ='"&curclassid&"' order by joindate desc  "
 	    else
 	       SQL="select top 10 * from userjoinclassinfo where  classid ='"&curclassid&"' and recordid not in (Select top "&Cstr(PageSize*(page-1))&" recordid from userjoinclassinfo where classid ='"&curclassid&"' order by joindate  ) order by joindate" 
            end if
            
            rs.open SQL,schooldb
            while not rs.eof 
	  %>          
        <tr> 
          <td width="300" height="20" valign="bottom"><%=rs("userid")%></td>
          <td width="210" valign="bottom" class="topic"><a href="deletecy.asp?pername=<%=rs("userid")%>">�����Ա</a></td>
        </tr>
       <%
            rs.movenext
            wend
            rs.close
       %> 
      </table>
      <br>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0" bgcolor="#C7E4F1">
        <form name="form1" method="post" action="popwindow_manage_deletecy.asp">
        <tr> 
          <td align="center" height="30">��<%=pagecount%>ҳ | 
            <a href="popwindow_manage_deletecy.asp?page=1">��һҳ</a> 
            | <a href="popwindow_manage_deletecy.asp?page=<%=page-1%>">��һҳ</a> 
            | <a href="popwindow_manage_deletecy.asp?page=<%=page+1%>">��һҳ</a> 
            | <a href="popwindow_manage_deletecy.asp?page=<%=pagecount%>">ĩҳ</a> | �� 
            <input type="text" name="page" size="2">
            ҳ 
            <input type="submit" name="Submit2" value="GO">
          </td>
        </tr>
        </form>
      </table>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>