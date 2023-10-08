
<!--#include file=globals.asp -->
<script language="javascript">
function selectclass(curid){

            if (curid!=0) {
                window.location.href="class_index.asp?selectclassid="+curid
        }
    }
function personalinfo(userid){


        window.open("popwindow_perinfo.asp?userID="+userid,"top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=580,height=400")

    }
function personallogout(){
               if (confirm("你确定要注销在这个班级的身份吗？"))
    {

        window.location.href="class_index.asp?act=del"
    }
    }
    function manage(){


        window.open("popwindow_manage.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=570,height=500")
            }
function PopWindowUpload(){


        window.open("PopWindowUpload.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=570,height=380")

    }

</script>
<%
set rs=server.createobject("adodb.recordset")
ClassID=Session("ClassID")
if  ClassID="" then response.redirect "../Error.asp?Info=请您先登陆，谢谢！"
if not UserID="" then
   SqlAdd=" and ToUserID='"&UserID&"'"
   '用户信息
    SQL = "select * from UserInfo where UserID='"&UserID&"'"
    'Response.Write sql
    rs.open SQL,DBParams
    if not rs.eof then
       RealName=rs("RealName")
    end if
    rs.close
    Act=2
else
   SqlAdd=" and ToUserID is null"
    '用户信息

    SQL = "select * from ClassInfo where ClassID='"&ClassID&"'"
    'Response.Write sql
    rs.open SQL,DBParams
    if not rs.eof then
       RealName=rs("ClassName")
    end if
    rs.close
    Act=1
end if

%>
<%

userid=session("userid")
 realname=session("realname")
set rs = createobject("ADODB.recordset")
selectclassid=trim(request("selectclassid"))
act=trim(request("act"))
if act="del" then
   sql="select * from classinfo where classid='"&session("classid")&"'"
   rs.open SQL,schooldb
   if not rs.eof then
      monitor=rs("monitor")
      monitor1=rs("monitor1")
   end if
   rs.close
   if userid=monitor then
      response.redirect "../error.asp?info=对不起您是本班的正班长，您不能注销自己的身份！"
   end if
   if userid= monitor1 then
      sql="select * from classinfo where classid='"&session("classid")&"'"
      rs.open SQL,schooldb,1,3
      if not rs.eof then
         rs("monitor1")=""
         rs.update
      end if
      rs.close
   end if
   sql="delete from userjoinclassinfo where userid='"&userid&"' and classid='"&session("classid")&"'"
   rs.open SQL,schooldb,1,3
   SQL = "select * from userjoinclassinfo where userid='"&userid&"' order by joindate"
   rs.open SQL,schooldb
   if not rs.eof then
      response.redirect "class_index.asp?selectclassid="&rs("classid")
   else
      Session.Abandon
      response.redirect "../success.asp?info=您已经注销了在该班级的身份，而且您现在没有加入任何一个班级，<a href='http://www.sun2003.com/school'>请点这里</a>,登陆后重新注册班级！"
   end if
   rs.close
end if
if selectclassid<>"" then
   sql="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&selectclassid&"'"
   rs.open SQL,schooldb
   if rs.eof then
      response.redirect "../error.asp?info=对不起，您不是本班成员无权进入！"
   end if
   rs.close
   session("classid")=selectclassid
   sql="select * from online where userid='"&userid&"'"
   rs.open SQL,schooldb,1,3
   if not rs.eof then
      rs("classid")=selectclassid

      rs.update
   end if
   rs.close
   sql="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&selectclassid&"'"
   rs.open SQL,schooldb,1,3
   if not rs.eof then
     if session("loginflag")="" then
        rs("logintimes")=rs("logintimes")+1
        session("loginflag")=true
     end if
     rs("lastlogintime")=now
     rs.update
   end if
   rs.close
end if


curclassid=Session("classid")
if userid="" then
  response.redirect "../error.asp?info=对不起，您已经掉线了，请重新进入！"
end if

set rss = createobject("ADODB.recordset")
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
  classname=rs("classname")
  monitor=rs("monitor")
  monitor1=rs("monitor1")
  tonggao=rs("tonggao")
   regdate=rs("regdate")
   enterdate=rs("enterdate")
end if
rs.close

schoolid=left(curclassid,8)
sc=schoolid
sql="select * from schoolinfo where schoolid='"&schoolid&"'"
rs.open SQL,schooldb
if not rs.eof then
   schoolname=rs("schoolname")
end if
rs.close


%>
<HTML><HEAD>
<meta http-equiv="Content-Language" content="zh-cn">
<TITLE>商务校友录 留言簿</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT language=javascript src="img/alumni.js"></SCRIPT>
<LINK href="img/alumni.css" type=text/css rel=stylesheet>
<script language="javascript">
function Del(URL)
   {
    if (confirm("你确定要删除吗？"))
       {
        //window.open("notebook.asp?curID="+curid+"&act=del","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=640,height=400")
       window.location.href=URL
       }
   }

</script>

<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=white leftMargin=0 topMargin=5 >
<CENTER>
<!--#include file=../top.htm-->
<TABLE cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR>
    <TD vAlign=top bgColor=#fbaf0c height=7 colspan="2" width="760">
    <table cellSpacing="0" cellPadding="0" width="760" border="0" height="1">
      <tr>
        <td width="141" bgcolor="#FFFFFF" height="1" valign="bottom">
        <img src="img/cr_a1.gif" width="133" height="6"></td>
        <td width="72" bgcolor="#FFFFFF" height="1">　</td>
        <td width="98" bgcolor="#FFFFFF" height="1">　</td>
        <td width="105" bgcolor="#FFFFFF" height="1">　</td>
        <td width="108" bgcolor="#FFFFFF" height="1">　</td>
        <td width="102" bgcolor="#FFFFFF" height="1">　</td>
        <td vAlign="top" width="134" bgcolor="#FFFFFF" height="1">　</td>
      </tr>
      <tr>
        <td width="141" height="32" background="img/cr_a31.gif">
        <img src="img/cr_a2.gif" border="0" width="141" height="32"></td>
        <td width="72" background="img/cr_a31.gif" height="32">
        　</td>
        <td width="98" height="32" background="img/cr_a31.gif">
        　</td>
        <td width="105" height="32" background="img/cr_a31.gif">
        　</td>
        <td width="108" height="32" background="img/cr_a31.gif">
        　</td>
        <td width="102" background="img/cr_a31.gif" height="32">
        　</td>
        <td vAlign="top" width="134" background="img/cr_a32.gif" height="32">
        <table height="26" cellSpacing="0" cellPadding="0" width="120" border="0">
          <tr>
            <td><a class="cnn" href="javascript:login()">
            <img src="img/cr_a22.gif" border="0" width="16" height="16"></a></td>
            <td class="cnn"><a class="cnn" href="javascript:login()">
            <font color="#ffffff">登录</font></a></td>
            <td width="12">
            <img src="img/cr_a23.gif" width="2" height="16"></td>
            <td><a class="cnn" href="../classlogout.asp">
            <img src="img/cr_a22.gif" border="0" width="16" height="16"></a></td>
            <td class="cnn"><a class="cnn" href="../classlogout.asp">
            <font color="#ffffff">注销</font></a></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </TD>
    </TR>
  <TR>
    <TD vAlign=top bgColor=#fbaf0c height=7 width="607"><IMG height=1 
      src="img/c.gif" width=1></TD>
    <TD align=middle width=157 bgColor=#FBAF0C><IMG height=1 
      src="img/c.gif" width=1></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=761 border=0>
  <TBODY>
  <TR>
    <TD class=bg1 style="BORDER-LEFT: #fbaf0c 7px solid" vAlign=top align=middle 
    width=123>
      <TABLE height=5 cellSpacing=0 cellPadding=0 width=100 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=1 src="img/c.gif" 
        width=1></TD></TR></TBODY></TABLE><LINK href="img/photo.css" 
      type=text/css rel=stylesheet>
      <TABLE height=22 cellSpacing=2 cellPadding=0 width=114 
      background=img/cla_bg3.gif border=0>
        <TBODY>
        <TR>
          <TD vAlign=bottom align=middle>我的校友录</TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=114 border=0 align=middle>
        <TBODY align=middle >
   
          &nbsp;&nbsp;您好:<font color="#FF6600"><%=userid%></font></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=114   height=362 border=0>
        <TBODY>
        <TR vAlign=top>
          <TD height="14" width="114" colspan="3" align="center">
           

                      </TD></TR>
        <TR vAlign=top>
          <TD height="7" width="114" colspan="3">
           

                      </TD></TR>
        <tr>
          <TD height="18" align="center" width="103" colspan="2">
           

                      ·个人设置</TD>
          <TD height="18" align="center" width="11">
           

                      </TD>
        </tr>
        <tr>
          <TD height="22" align="center" valign="top" width="114" colspan="3">
           

                      <a class="cla1" href="popwindow_modifyinfo.asp" style="text-decoration: none">
                      <font color="#FF6600">个人信息</font></a></TD>
        </tr>
        <tr>
          <TD height="18" align="center" width="103" colspan="2">
           

                      ·班级信息</TD>
          <TD height="18" align="center" valign="top" width="11">
           

                      </TD>
        </tr>
        <tr>
          <TD height="112" align="center" valign="top" width="114" colspan="3">
           

                      <a class="cla1" href="ClassBBSList.asp">
                      <font color="#FF6600">班级留言</font></a><br>
                      <a class="cla1" href="class_phonebook.asp">
                      <font color="#FF6600">成员地址</font></a><br>
                      <a class=cla1 href="PopWindowEmailList.asp">
                      <font color="#FF6600">邮件列表</font></a><br>
                      <a class="cla1" href="ClassPhotoList.asp">
                      <font color="#FF6600">班级相册</font></a><br>
                      <a class="cla1" href="javascript:PopWindowUpload();">
                      <font color="#FF6600">上载照片</font></a><br>
                      <a class="cla1" href="tj.asp" style="text-decoration: none">
                      <font color="#FF6600">访问统计</font></a></TD>
        </tr>
        <tr>
          <TD height="18" align="center" width="91">
           

                      ·管理员</TD>
          <TD height="18" align="center" width="12">
           

                      </TD>
          <TD height="18" align="center" width="11">
           

                      </TD>
        </tr>
        <tr>
          <TD height="22" align="center" valign="top" width="114" colspan="3">
           

                <a class="cla1" href="javascript:manage();">                      <font color="#FF6600">
                班级管理</font></a></TD>
        </tr>
        <tr>
          <TD height="17" align="center" width="103" colspan="2">
           

                      ·相关信息</TD>
          <TD height="17" align="center" valign="top" width="11">
           

                      </TD>
        </tr>
        <tr>
          <TD height="114" align="center" valign="top" width="114" colspan="3">
           

                      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="103">
                        <tr>
                          <td width="29%" height="103">　</td>
                          <td width="71%" height="103" valign="top">
                          <a class="cla1" href="../register.htm" style="text-decoration: none">
                          <font color="#FF6600">新用户注册</font></a><br>
                          <a class="cla1" href="../schoollist.asp" style="text-decoration: none">
                          <font color="#FF6600">浏览学校</font></a><br>
                          <a class="cla1" href="../find_class1.asp" style="text-decoration: none">
                          <font color="#FF6600">注册新班级</font></a></td>
                        </tr>
                      </table>
          </TD>
        </tr>
        </TBODY></TABLE></TD>
    <TD vAlign=top align=middle width=620 
    background=img/cla_bg2.gif>
      <TABLE cellSpacing=0 cellPadding=0 width=607 border=0>
        <TBODY>
        <TR>
          <TD class=cla1 width=380 height=24>&nbsp;&nbsp;校友录首页 &gt;&gt;<a href="../classlist.asp?schoolid=<%=schoolid%>" class=6 target="_blank"><%=schoolname%></a> 
          &gt;&gt;<%=classname%></TD>
          <TD align=right width=607></TD></TR>
        <TR align=middle>
          <TD colSpan=2 width="607"><IMG height=3 src="img/cla_a1.gif" 
            width=457></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=1 cellPadding=0 width=609 bgColor=#d2d2d2 border=0>
        <TBODY>
        <TR>
          <TD bgColor=#ffffff width="607">
            <div align="center">
              <center>
            <TABLE class=cr4 cellSpacing=0 cellPadding=3 width=623 
            border=0 style="border-collapse: collapse" bordercolor="#111111">
              <TBODY>
              <TR align=middle>
                <TD width="49"><a class="down" href="ClassBBSList.asp"><IMG 
                  height=26 src="img/cla_a40.gif" width=36 
                  border=0></a><BR><a class="cla1" href="ClassBBSList.asp">班级留言</a><BR></TD>
                <TD width="8"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0><BR></TD>
                <TD width="49"><a class="down" href="class_phonebook.asp"><IMG 
                  height=26 src="img/cla_a41.gif" width=36 
                  border=0></a><BR><a class="cla1" href="class_phonebook.asp">
                成员地址 
                </a> 
                </TD>
                <TD width="6"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD width="51"><a class="down" href="popwindow_search.html"><IMG 
                  height=26 src="img/cla_a42.gif" width=36 
                  border=0></a><BR><a class="cla1" href="popwindow_search.html">
                校友查询</a> 
                </TD>
                <TD width="5"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD width="48"><a class="down" href="PopWindowEmailList.asp"><IMG 
                  height=26 src="img/cla_a43.gif" width=36 
                  border=0></a><BR><a class=cla1 href="PopWindowEmailList.asp">
                邮件列表</a></TD>
                <TD width="5"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD width="48">
                <a class="down" href="javascript:manage();"><IMG 
                  height=26 src="img/cla_a44.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" href="javascript:manage();">班级管理</a></TD>
                <TD width="4">
                <IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD width="284">
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                  <tr>
                <TD align="center"><a class="down" href="ClassPhotoList.asp"><IMG 
                  height=26 src="img/cla_a46.gif" width=36 
                  border=0></a><BR><a class="cla1" href="ClassPhotoList.asp">
                班级相册 
                </a> 
                </TD>
                <TD align="center"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD align="center"><a class="down" href="javascript:PopWindowUpload();"><IMG 
                  height=26 src="img/cla_a45.gif" width=36 
                  border=0></a><BR><a class="cla1" href="javascript:PopWindowUpload();">
                上载照片</a> 
                </TD>
                <TD align="center"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD align="center"><a class="down" href="popwindow_modifyinfo.asp"><IMG 
                  height=26 src="img/cla_a47.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" href="popwindow_modifyinfo.asp">修改资料</a> 
                </TD>
                <TD align="center"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD align="center">
                <a class="down" href="popwindow_modifyinfo.asp"><IMG 
                  height=26 src="img/cla_a48.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" href="popwindow_modifyinfo.asp">个人信息</a></TD>
                <TD align="center"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD align="center"><a class="down" href="../find_class1.asp"><IMG height=26 
                  src="img/cla_a49.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" id="t0" target="_blank" href="../find_class1.asp">
                注册新班</a></TD>
                  </tr>
                </table>
                </TD></TR>
              </TBODY></TABLE></center>
            </div>
          </TD></TR>
        </TBODY></TABLE>
      <TABLE cellSpacing=3 cellPadding=0 width=453 border=0>
        <TBODY>
        <TR>
          <TD align=right width=76>　</TD>
          <TD width="368">
          　</TD></TR>
        </TBODY></TABLE>
      <TABLE cellSpacing=3 cellPadding=0 width=577 border=0 style="border-collapse: collapse" bordercolor="#111111" height="76">
        <TBODY>
        <tr>
          <TD bgColor=#C8F4FF height=20 width="576">&nbsp;&nbsp; <strong>
          <font color="#246F00">班级留言</font></strong></TD>
        </tr>
        <TR>
          <TD height=21 width="576">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" align="center">
   
<%

       PageSize=15
    rs.open "Select count(*) as c from message where ClassID='"&ClassID&"' and Deleted=0 "&SqlAdd,DBParams,1,3
    PageCount=CInt(rs("c")/PageSize+0.5)
    rs.Close
    Page = CLng(Request("Page"))
    if Page=""  or isnull(Page) then  page=1
    If Page < 1 Then    Page = 1
    If Page > PageCount Then  Page = PageCount
%>
      <p>
    
<%


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
%> </p>
      <table width="573" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="22" width="17" bgcolor="#E1E1E1"><img src="../school_images/topic_inco1.gif" width="15" height="15"></td>
          <td height="22" width="556" class="topic2" bgcolor="#E1E1E1">标&nbsp;&nbsp;题：<%=rs("Title")%></td>
        </tr>
        <tr>
          <td height="22" valign="middle" width="17"><img src="../school_images/mood<%=rs("HeadPic")%>.gif" width="15" height="15" border="0" alt="删除"></td>
          <td height="22" valign="middle" width="556" class="topic2">留言人：<%=rs("UserID")%></td>
        </tr>
        <tr>
          <td height="4" valign="top" colspan="2" width="573"></td>
        </tr>
        <tr>
          <td class="topic" valign="top" colspan="2" width="573"><%=rs("Comment")%></td>
        </tr>
        <tr valign="bottom" align="right">
          <td class="di" colspan="2" height="20" width="573"> 留言于：<%=rs("RegTime")%>&nbsp;&nbsp;</td>
        </tr>
      </table>
      <p>
      </p>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td bgcolor="#B4C7D4" height="1"></td>
        </tr>
      </table>
      <p>
    
<%
      rs.MoveNext
   wend
   rs.close
%> </p>
      <table width="572" border="0" cellspacing="0" cellpadding="0" bgcolor="#C7E4F1" height="24">
        <tr>
          <td align="right" height="24" width="572" bgcolor="#E2F9FF">共<%=PageCount%>页 | <a href="ClassBBSList.asp?Page=1&CLassID=<%=ClassID%>&UserID=<%=UserID%>">第一页</a>
            | <a href="ClassBBSList.asp?Page=<%=Page-1%>&CLassID=<%=ClassID%>&UserID=<%=UserID%>">上一页</a>
            | <a href="ClassBBSList.asp?Page=<%=Page+1%>&CLassID=<%=ClassID%>&UserID=<%=UserID%>">下一页</a>
            | <a href="ClassBBSList.asp?Page=<%=PageCount%>&CLassID=<%=ClassID%>&UserID=<%=UserID%>">末页</a>
            | </td>
        </tr>
      </table>
      <p>
    
      </p>
      <table width="573" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="20" class="bt" width="100"><a name="WYLY"></a>我要留言 |</td>
          <td height="20" class="topic" width="473" align="right" valign="top">　</td>
        </tr>
        <tr>
          <td bgcolor="#B4C7D4" height="1" colspan="2" width="573"></td>
        </tr>
      </table>

      <table width="570" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#EBEBEB">
         <form method="post" action="ClassBBSPost.asp">

          <tr>
            <td height="30" class="topic2" width="80" align="right">标&nbsp;&nbsp;题：
            </td>
            <td height="30" valign="middle" class="topic2" width="480">
              <input type="text" name="Title" size="50" value="大家好">
            </td>
          </tr>
          <tr>
            <td height="30" class="topic2" align="right">留言人： </td>
            <td height="30" valign="middle" class="topic2">
              <input type="text" name="UserID" value="<%=Session("RealName")%>" size="20">
            </td>
          </tr>
          <tr>
            <td class="topic2" align="right">内&nbsp;&nbsp;容：</td>
            <td class="topic" valign="top">
              <textarea name="Comment" cols="55" rows="10"></textarea>
            </td>
          </tr>
          <tr valign="middle" align="left">
            <td class="topic" height="30" valign="bottom" colspan="2">&nbsp;&nbsp;&nbsp; 选择表情：</td>
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
     
      <p>　</td>
            </tr>
          </table>
          </TD></TR></TBODY></TABLE>
      <TABLE class=cr4 cellSpacing=0 cellPadding=5 width=430 align=center 
      border=0>
        <TBODY>
        <TR>
          
</TD></TR></TBODY></TABLE>
      </TD>
    <TD style="BORDER-LEFT: #fbaf0c 7px solid" vAlign=top align=middle width=4 
    bgColor=#FBAF0C>
      　</TD></TR></TBODY></TABLE>
<TABLE height=1 cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR>
    <TD vAlign=top bgColor=#fbaf0c height="1"><IMG height=1 src="img/c.gif" 
      width=1></TD>
    <TD align=middle width=157 bgColor=#FBAF0C height="1">
    　</TD></TR></TBODY></TABLE>
</CENTER>
<!--#include file=../end.htm--></BODY></HTML>