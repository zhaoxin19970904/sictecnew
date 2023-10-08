<!--#include file=globals.asp -->
<%
set rs = server.createobject("ADODB.recordset")

ClassID=Session("ClassID")
if ClassID="" then ClassID="GoodLuck"
UserID=Trim(Session("UserID"))
if  UserID="" then response.redirect "../Error.asp?Info=请您先登陆，谢谢！"
   SqlAdd=" and ToUserID='"&UserID&"'"   
   '用户信息
    SQL = "select * from UserInfo where UserID='"&UserID&"'"
    'Response.Write sql
    rs.open SQL,DBParams
    if not rs.eof then   
       RealName=rs("RealName")
    end if
    rs.close
%> 
<HTML>
<HEAD>
<title>商务同学录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<meta name="web_designner" content="孙梦昕">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
<script language="javascript">
function Del(URL)
   {
	if (confirm("你确定要删除吗？"))
	   {
	    //window.open("notebook.asp?curID="+curid+"&act=del","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=640,height=400")
	   window.location.href=URL
	   }
   }
function personalinfo(userid){			
		window.open("popwindow_perinfo.asp?userID="+userid,"top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=580,height=400")
	}
</script>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#52707D"><img src="../school_images/poplogo.gif" width="108" height="34"></td>
  </tr>
  <tr>
    <td bgcolor="#52707D" background="../school_images/popwindow_03.gif">&nbsp;</td>
  </tr>
</table>
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"> <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" class="topic" width="100" valign="top">私人留言 |</td>
          <td height="20" class="topic" width="410" align="right" valign="top">&nbsp;</td>
        </tr>
        <tr> 
          <td bgcolor="#B4C7D4" height="1" colspan="2"></td>
        </tr>
      </table>
      <br>
      <%

    PageSize=10
	rs.open "Select count(*) as c from message where ClassID='"&ClassID&"' and Deleted=0 "&SqlAdd,DBParams,1,3
	PageCount=CInt(rs("c")/PageSize+0.5)
	rs.Close		
%> 
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="28" class="bt" width="100"><%=RealName%>留言簿</td>
          <td height="28" class="topic" width="410" align="right" valign="top">共<%=PageCount%>页 
            | <a href="popwindow_perliuyan.asp?Page=1&amp;CLassID=<%=ClassID%>&amp;UserID=<%=UserID%>">第一页</a> 
            | <a href="popwindow_perliuyan.asp?Page=<%=Page-1%>&amp;CLassID=<%=ClassID%>&amp;UserID=<%=UserID%>">上一页</a> 
            | <a href="popwindow_perliuyan.asp?Page=<%=Page+1%>&amp;CLassID=<%=ClassID%>&amp;UserID=<%=UserID%>">下一页</a> 
            | <a href="popwindow_perliuyan.asp?Page=<%=PageCount%>&amp;CLassID=<%=ClassID%>&amp;UserID=<%=UserID%>">末页</a> 
            | </td>
        </tr>
        <tr> 
          <td bgcolor="#B4C7D4" height="1" colspan="2"></td>
        </tr>
      </table>
      <br>
      <%    
    Page = CLng(Request("Page"))
    if Page=""  or isnull(Page) then  page=1
    If Page < 1 Then    Page = 1
    If Page > PageCount Then  Page = PageCount

        if page=1 or page=0 then
		SQL="select top "&PageSize&" * from message where ClassID='"&ClassID&"' and Deleted=0 "&SqlAdd&" order by MsgID DESC"
	else
		SQL="select  top "&PageSize&" * from message where ClassID='"&ClassID&"' and Deleted=0 "&SqlAdd&" and "
		SQL=SQL & " MsgID not in (Select top "&Cstr(PageSize*(page-1))&" MsgID from message where "
		SQl=SQL & " ClassID='"&ClassID&"' and Deleted=0 "&SqlAdd&" order by MsgID DESC)"
		SQL=SQL & " order by MsgID DESC"
	end if
	'Response.Write sql&page
	rs.Open SQL,DBParams
   while not rs.eof	
%> 
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="22" valign="middle" width="15"><img src="../school_images/topic_inco1.gif" width="15" height="15"></td>
          <td height="22" valign="middle" width="495" class="topic2">标&nbsp;&nbsp;题：<%=rs("Title")%></td>
        </tr>
        <tr> 
          <td height="22" valign="middle" width="15"><a href="ClassDelAction.asp?Act=2&amp;ClassID=<%=ClassID%>&amp;MsgID=<%=rs("MsgID")%>&amp;UserID=<%=UserID%>"><img src="../school_images/mood<%=rs("HeadPic")%>.gif" width="15" height="15" border="0" alt="删除"></a></td>
          <td height="22" valign="middle" width="495" class="topic2">留言人：<a href="javascript:personalinfo('<%=rs("UserID")%>');"><%=rs("UserID")%></a></td>
        </tr>
        <tr> 
          <td height="4" valign="top" colspan="2"></td>
        </tr>
        <tr> 
          <td class="topic" valign="top" colspan="2"><%=rs("Comment")%></td>
        </tr>
        <tr valign="bottom" align="right"> 
          <td class="di" colspan="2" height="20">留言于：<%=rs("RegTime")%>&nbsp;&nbsp;<a href="javascript:Del('ClassDelAction.asp?Act=2&amp;MsgID=<%=rs("MsgID")%>&amp;UserID=<%=UserID%>')">删除</a></td>
        </tr>
      </table>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#B4C7D4" height="1"></td>
        </tr>
      </table>
      <br>
<%
      rs("Hits")=rs("Hits")+1
      rs.update
      rs.MoveNext
   wend
rs.close
%> 
<br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" class="bt" colspan="2">给校友私人留言 |</td>
        </tr>
        <tr> 
          <td bgcolor="#B4C7D4" height="1" colspan="2"></td>
        </tr>
      </table>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <form method="post" action="ClassBBSPost.asp">
          <tr> 
            <td height="30" valign="middle" class="topic2" width="80">校友： </td>
            <td height="30" valign="middle" class="topic2" width="480"> 
              <select name="ToUserID">
                <option value="0" selected>选择给谁私人留言</option>
<%
sql="select U2.UserID,U2.RealName from UserJoinClassInfo as U1,UserInfo as U2  where U1.UserID=U2.UserID  and  U1.ClassID='"&ClassID&"'"
response.write sql
rs.Open SQL,DBParams
while not rs.eof
   response.write "<option value="&rs("UserID")&">"&rs("RealName")&"</option>"   
   rs.MoveNext
wend
set rs = nothing
%> 

              </select>
            </td>
          </tr>
          <tr> 
            <td height="30" valign="middle" class="topic2" width="80">标&nbsp;&nbsp;题： 
            </td>
            <td height="30" valign="middle" class="topic2" width="480"> 
              <input type="text" name="Title" size="60" value="大家好">
            </td>
          </tr>
          <tr> 
            <td height="30" valign="middle" class="topic2">留言人： </td>
            <td height="30" valign="middle" class="topic2"> 
              <input type="text" name="UserID" value="<%=Session("UserID")%>">
            </td>
          </tr>
          <tr> 
            <td class="topic2" valign="middle">内&nbsp;&nbsp;容：</td>
            <td class="topic" valign="top"> 
              <textarea name="Comment" cols="65" rows="10"></textarea>
            </td>
          </tr>
          <tr valign="middle" align="left"> 
            <td class="topic" height="30" valign="bottom" colspan="2">选择表情：</td>
          </tr>
          <tr valign="bottom" align="center"> 
            <td class="di" valign="middle" height="35" align="center" colspan="2"> 
              <input type="radio" name="HeadPic" value="0" checked>
              <img src="../school_images/mood0.gif" width="15" height="15"> &nbsp; 
              &nbsp; 
              <input type="radio" name="HeadPic" value="1">
              <img src="../school_images/mood1.gif" width="15" height="15">&nbsp; 
              &nbsp; 
              <input type="radio" name="HeadPic" value="2">
              <img src="../school_images/mood2.gif" width="15" height="15">&nbsp; 
              &nbsp; 
              <input type="radio" name="HeadPic" value="3">
              <img src="../school_images/mood3.gif" width="15" height="15">&nbsp;&nbsp; 
              <input type="radio" name="HeadPic" value="4">
              <img src="../school_images/mood4.gif" width="15" height="15">&nbsp;&nbsp; 
              <input type="radio" name="HeadPic" value="5">
              <img src="../school_images/mood5.gif" width="15" height="15">&nbsp;&nbsp; 
              <input type="radio" name="HeadPic" value="6">
              <img src="../school_images/mood6.gif" width="15" height="15">&nbsp;&nbsp; 
              <input type="radio" name="HeadPic" value="7">
              <img src="../school_images/mood7.gif" width="15" height="15"> </td>
          </tr>
          <tr valign="bottom" align="center"> 
            <td class="di" height="35" colspan="2"> 
              <input type="hidden" name="ClassID" value="<%=ClassID%>">
              <input type="submit" name="Submit" value="提交">
              &nbsp;&nbsp; 
              <input type="reset" name="Submit3" value="重写">
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