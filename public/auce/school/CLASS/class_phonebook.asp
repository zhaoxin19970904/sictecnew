
<%@ Language=VBScript %>
<%Response.buffer=true%>
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
sql="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
   logintimes=rs("logintimes")
   userstatus=rs("userstatus")

end if
rs.close
sql="select count(*) as a from userjoinclassinfo where userid='"&userid&"'"
rs.open SQL,schooldb,1,3
joinnum=rs("a")
rs.close
sql="select count(*) as b from userjoinclassinfo where classid='"&curclassid&"' and userstatus='成员'"
rs.open SQL,schooldb,1,3
membernum=rs("b")
rs.close
sql="select count(*) as c from userjoinclassinfo where classid='"&curclassid&"' and userstatus='教师'"
rs.open SQL,schooldb,1,3
teachernum=rs("c")
rs.close
sql= "Select count(*) as c from message where ClassID='"&curclassid&"' and Deleted=0 and Hits=0 and ToUserID='"&userid&"'"
rs.open SQL,schooldb,1,3
MsgNum=rs("c")
rs.close
%>
<%
set rs=server.createobject("adodb.recordset")
ClassID=Session("ClassID")
%>
<HTML><HEAD>
<meta http-equiv="Content-Language" content="zh-cn">
<TITLE>商务校友录 <%=schoolname%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT language=javascript src="img/alumni.js"></SCRIPT>
<LINK href="img/alumni.css" type=text/css rel=stylesheet>


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
                      <a class="cla1" href="javascript:PopWindowUpload();">                      <font color="#FF6600">上载照片</font></a><br>
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
          <TD align=right width=227></TD></TR>
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
                  border=0></a><BR><a class="cla1" href="javascript:PopWindowUpload();">                上载照片</a> 
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
      <TABLE cellSpacing=3 cellPadding=0 width=600 border=0 style="border-collapse: collapse" bordercolor="#111111" height="79">
        <TBODY>
        <TR>
          <TD height=11 width="594" align="center">
          <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CECECE" width="590">
            <tr>
              <td align="middle" bgColor="#f7f7f7" width="588">
              <table height="28" cellSpacing="0" cellPadding="0" border="0">
                <tr>
                  <td width="31">
                  <img height="22" src="http://images.sohu.com/cs/sms/alumni3/images/clam_071.gif" width="24">
                  </td>
                  <td><strong><%=classname%>成员地址</strong></td>
                </tr>
              </table>
              </td>
            </tr>
          </table>
          </TD></TR>
        <TR>
          <TD height=12 width="594"></TD></TR>
        <tr>
          <TD bgColor=#C8F4FF height=23 width="594">&nbsp;&nbsp;
          <font color="#246f00"><strong>成员</strong></font><strong><font color="#246F00">详细信息</font></strong></TD>
        </tr>
        <TR>
          <TD height=21 width="594">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%">
      <br>    <%
            
            if Request("Page")<>"" then 
       		Page = CLng(Request("Page"))
    	    end if 
            If Page < 1 Then 
                Page = 1
            end if
            PageSize=5
            sql="select count(*) as a from userjoinclassinfo where classid ='"&curclassid&"' and userstatus<>'教师'"  
            rs.open SQL,schooldb,1,3
            count=rs("a")
            PageCount=CInt(rs("a")/PageSize+0.5)
	    rs.Close
       %>
      <%
            if page=1 then
	       SQL="select top 5 * from userjoinclassinfo where classid ='"&curclassid&"' and userstatus<>'教师' order by userid  "
 	    else
 	       SQL="select top 5 * from userjoinclassinfo where classid ='"&curclassid&"' and userstatus<>'教师' and userid not in (Select top "&Cstr(PageSize*(page-1))&" userid from userjoinclassinfo where classid ='"&curclassid&"' and userstatus<>'教师' order by userid  ) order by userid" 
            end if
            
            rs.open SQL,schooldb
            while not rs.eof 
            sql1="select * from userinfo where userid='"&rs("userid")&"'"
            rss.open SQL1,schooldb
            if not rss.eof then
               realname=rss("realname")
            end if
            rss.close   
            sql1="select * from usercommunicationinfo where userid='"&rs("userid")&"'"
            rss.open SQL1,schooldb
            if not rss.eof then
      %>      
      <table width="510" cellspacing="1" cellpadding="1" align="center">
        <tr> 
          <td colspan="2" height="20" class="topic">用户名：<%=rss("userid")%></td>
          <td height="20" class="topic">校友姓名：<%=realname%></td>
        </tr>
        <tr class="digi"> 
          <%if rss("isqqshow")=1 then%>
          <td colspan="3" height="20">OICQ/ICQ号码：<%=rss("qq")%></td>
          <%else%>
          <td colspan="3" height="20">OICQ/ICQ号码：******</td>
          <%end if%>
        </tr>
        <tr class="digi"> 
          <%if rss("istelephoneshow")=1 then%>
          <td width="200" height="20">固定电话：<%=rss("telephone")%></td>
          <%else%>
          <td width="200" height="20">固定电话：******</td>
          <%end if%>
          <%if rss("ismobileshow")=1 then%>
          <td width="200" height="20">移动电话：<%=rss("mobile")%></td>
          <%else%>
          <td width="200" height="20">移动电话：******</td>
          <%end if%>
          <%if rss("isbpshow")=1 then%>
          <td width="200" height="20">传呼：<%=rss("bp")%></td>
          <%else%>
          <td width="200" height="20">传呼：******</td>
          <%end if%>
          
        </tr>
        <tr> 
          <%if rss("iscashow")=1 then%>
          <td colspan="3" height="20" class="topic">通信地址：<%=rss("communicationaddr")%></td>
          <%else%>
           <td colspan="3" height="20" class="topic">通信地址：**********</td>
          <%end if%>
        </tr>
        <tr> 
           <%if rss("ishashow")=1 then%>
          <td colspan="3" height="20" class="topic">居住地址：<%=rss("homeaddr")%></td>
          <%else%>
           <td colspan="3" height="20" class="topic">居住地址：**********</td>
          <%end if%>
          
        </tr>
        <tr> 
           <%if rss("iswsshow")=1 then%>
          <td colspan="3" height="20" class="topic">工作单位：<%=rss("workshop")%></td>
          <%else%>
           <td colspan="3" height="20" class="topic">工作单位：**********</td>
          <%end if%>
          
        </tr>
        <tr> 
          
          <%if rss("isemailshow")=1 then%>
          <td colspan="3" height="20" class="digi">电子邮件：<a href="mailto:<%=rss("email")%>" class=4><%=rss("email")%></a></td>
          <%else%>
           <td colspan="3" height="20" class="digi">电子邮件：**********</td>
          <%end if%>
          
        </tr>
      </table>
      <hr width="510" size="1" noshade align="center">
     <%end if
       rss.close
        rs.movenext
           wend
           rs.close
         %>
              <p>  

              </p>
      <div align="center">
        <center>
      <table width="590" border="0" cellspacing="0" cellpadding="0" bgcolor="#F7F7F7" style="border-collapse: collapse" bordercolor="#111111">
        <form name="form4" method="post" action="class_phonebook.asp">
        <tr> 
          <td height="28" class="topic" width="410" align="right" valign="top">共<%=pagecount%>页 
            | <a href="class_phonebook.asp?page=1">第一页</a> 
            | <a href="class_phonebook.asp?page=<%=page-1%>">上一页</a> 
            | <a href="class_phonebook.asp?page=<%=page+1%>">下一页</a> 
            | <a href="class_phonebook.asp?page=<%=pagecount%>">末页</a> | 到 
            <input type="text" name="page" size="2">
            页 
            <input type="submit" name="Submit22" value="GO">
          </td>
        </tr>
        </form>
      </table>
              </center>
      </div>
              </td>
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