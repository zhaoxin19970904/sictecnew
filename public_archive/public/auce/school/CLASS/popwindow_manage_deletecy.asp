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
if userid<>monitor or userid<>monitor1 then
   response.redirect "../error.asp?info=对不起，您不是本班管理员无权进入！"
end if  

%> 
<HTML>
<HEAD>
<title>班级管理--驱逐成员</title>
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
          <td>驱逐成员---</td>
        </tr>
      </table>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" class="topic" width="300"><font color="#CC0000">成员</font></td>
          <td height="20" class="digi" width="210" align="left" valign="middle">　</td>
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
          <td width="210" valign="bottom" class="topic"><a href="deletecy.asp?pername=<%=rs("userid")%>">驱逐成员</a></td>
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
          <td align="center" height="30">共<%=pagecount%>页 | 
            <a href="popwindow_manage_deletecy.asp?page=1">第一页</a> 
            | <a href="popwindow_manage_deletecy.asp?page=<%=page-1%>">上一页</a> 
            | <a href="popwindow_manage_deletecy.asp?page=<%=page+1%>">下一页</a> 
            | <a href="popwindow_manage_deletecy.asp?page=<%=pagecount%>">末页</a> | 到 
            <input type="text" name="page" size="2">
            页 
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