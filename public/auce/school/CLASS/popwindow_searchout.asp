<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
realname=trim(request.form("realname"))
schoolname=trim(request.form("schoolname"))
classname=trim(request.form("classname"))
enterdate=trim(request.form("enterdate"))
if (classname<>"" or enterdate<>"") and schoolname="" then
   response.redirect "../error.asp?info=对不起，您在查找的同学时如果填写了所在班级或入学年份，则一定要填写所在学校的名称 ！"
end if
if realname="" then
   response.redirect "../error.asp?info=对不起，您要查找的同学的姓名不能为空 ！"
end if

%>
<HTML>
<HEAD>
<title>商务校友录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0"><!--#include file=../top.htm-->
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="top"> <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="20" align="left" valign="top">下面是根据您填写的信息搜索到的结果：</td>
        </tr>
      </table>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" class="topic" width="100"><font color="#CC0000">校友姓名</font></td>
          <td height="20" class="digi" width="200" align="left" valign="middle"><font color="#CC0000">EMAIL</font></td>
          <td height="20" class="topic" width="210" align="left" valign="middle"><font color="#CC0000">所在班级</font></td>
        </tr>
        <tr> 
          <td bgcolor="#B4C7D4" height="1" colspan="3"></td>
        </tr>
      </table>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
       <%
        set rs = createobject("ADODB.recordset")
        set rss = createobject("ADODB.recordset")
        set rrs = createobject("ADODB.recordset")
        set result = createobject("ADODB.recordset")
   
        sql="select * from userinfo where realname='"&realname&"'"
        
        rs.open SQL,schooldb
        if rs.eof then
           response.write "<font color=red>对不起，没有您要找的同学</font>"
           response.end
        end if 
        while not rs.eof 
           sql1="select * from usercommunicationinfo where userid='"&rs("userid")&"'"
           rss.open sql1,schooldb
           if not rss.eof then
              email=rss("email")
           end if   
           rss.close
           temp=""
           if schoolname<>"" then
              if classname<>"" then
                 temp=" and classname='"&classname&"'"
              end if
              if enterdate<>"0" then
                 temp=temp&" and enterdate='"&enterdate&"'"
              end if      
           sql1="select * from schoolinfo where schoolname='"&schoolname&"'"
           rss.open sql1,schooldb
           if rss.eof then
              response.write "<font color=red>对不起，没有您要找的同学</font>"
              response.end
           else
              sql2="select * from classinfo where schoolid='"&rss("schoolid")&"'"&temp
              rrs.open sql2,schooldb
              
              if rrs.eof then
                 response.write "<font color=red>对不起，没有您要找的同学</font>"
                 response.end
              end if
              while not rrs.eof
                 sql3="select * from userjoinclassinfo where userid='"&rs("userid")&"' and classid='"&rrs("classid")&"'"
                 result.open sql3,schooldb          
                 while not result.eof 
                     if classname="" then
                        classname=rrs("classname")
                     end if   
       %>          
        <tr valign="bottom"> 
          <td width="100" height="20"><%=realname%></td>
          <td width="200"><a href="mailto:<%=email%>" class=4><%=email%></a></td>
          <td width="210"><a href="../classlist.asp?schoolid=<%=rss("schoolid")%>"><%=schoolname%></a>--<a href="class_index1.asp?classid=<%=rrs("classid")%>"><%=classname%></a></td>
        </tr>
        <%result.movenext
          wend 
          result.close
          rrs.movenext
          wend
          rrs.close
          end if
          rss.close
          else
          sql1="select * from userjoinclassinfo where userid='"&rs("userid")&"'"
        
          rss.open sql1,schooldb
          while not rss.eof 
                sql2="select * from classinfo where classid='"&rss("classid")&"'"
                rrs.open sql2,schooldb
                if not rrs.eof then
                   classname=rrs("classname")
                end if
                rrs.close
                
                schoolid=left(rss("classid"),8)
                sql2="select * from schoolinfo where schoolid='"&schoolid&"'"
                rrs.open sql2,schooldb
                if not rrs.eof then
                   schoolname=rrs("schoolname")
                end if
                rrs.close 
                  
         %>
          <tr valign="bottom"> 
          <td width="100" height="20"><%=realname%></td>
          <td width="200"><a href="mailto:<%=email%>" class=4><%=email%></a></td>
          <td width="210"><a href="../classlist.asp?schoolid=<%=schoolid%>"><%=schoolname%></a>--<a href="class_index1.asp?classid=<%=rss("classid")%>"><%=classname%></a></td>
        </tr>
        <%
         rss.movenext
         wend
         rss.close
         end if
         rs.movenext
         wend
         rs.close
        %> 
      </table>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>