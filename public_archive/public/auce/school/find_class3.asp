<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%userid=session("userid")
provinceid=trim(request.form("provinceid"))
areaid=trim(request("areaid"))
schooltypeid=int(trim(request("schooltypeid")))
keyword=trim(request("keyword"))
if keyword="" then
  response.redirect "error.asp?info=请输入学校名的关键字！"
end if
%>
<html>
<head>
<title>商务校友录-加入班级</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<STYLE type=text/css>
</STYLE>
<LINK href="scss.css" rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="5"><CENTER>
<!--#include file=top2.htm-->

<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" valign="top" bgcolor="#F6F6F6" width="760"> <br>
      <br>


      <%
            if Request("Page")<>"" then
               Page = CLng(Request("Page"))
            end if
            If Page < 1 Then
                Page = 1
            end if
       %>
      <table width="600" border="0" cellspacing="0" cellpadding="0">
        <tr align="left" valign="top">
          <td height="26" width="20" valign="middle"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
          <%
                    PageSize=5
                    set rs=Server.CreateObject("ADODB.RecordSet")
                    sql="select count(*) as a from schoolinfo where schooltypeid="&schooltypeid&" and schoolid like '"&areaid&"%' and schoolname like '%"&keyword&"%'"
                    rs.open SQL,schooldb,1,3
                    count=rs("a")
                    PageCount=CInt(rs("a")/PageSize+0.5)
                rs.Close
      %>
          <td height="20" valign="middle">根据您的关键字搜索到 <%=count%> 所学校，请选择加入：</td>
        </tr>
      </table>
      <table width="600" border="0" cellspacing="1" cellpadding="1">
        <form name="form1" method="post" action="find_class4.asp">
          <%
            if page=1 then
           SQL="select top 5 * from schoolinfo where schooltypeid="&schooltypeid&" and schoolid like '"&areaid&"%' and schoolname like '%"&keyword&"%' order by schoolid  "
         else
            SQL="select top 5 * from schoolinfo where schooltypeid="&schooltypeid&" and schoolid like '"&areaid&"%' and schoolname like '%"&keyword&"%' and schoolid not in (Select top "&Cstr(PageSize*(page-1))&" schoolid from schoolinfo where schooltypeid="&schooltypeid&" and schoolid like '"&areaid&"%' and schoolname like '%"&keyword&"%' order by schoolid  ) order by schoolid"
            end if

            rs.open SQL,schooldb
            while not rs.eof
            schoolname=replace(rs("schoolname"),keyword,"<font color=red>"&keyword&"</font>")
          %>
          <tr>
            <td height="28" bgcolor="#EBEBEB">
              <input type="radio" name="schoolid" value="<%=rs("schoolid")%>">
              <%=schoolname%></td>
          </tr>
          <%
            rs.movenext
            wend
            rs.close
          %>
          <tr>
            <td height="28" bgcolor="#CCCCCC" align="center" class="topic">当前页数<%=page%>/<%=pagecount%>页
              | <a href="find_class3.asp?page=1&&areaid=<%=areaid%>&schooltypeid=<%=schooltypeid%>&keyword=<%=keyword%>">第一页</a>
              | <a href="find_class3.asp?page=<%=page-1%>&areaid=<%=areaid%>&schooltypeid=<%=schooltypeid%>&keyword=<%=keyword%>">上一页</a>
              | <a href="find_class3.asp?page=<%=page+1%>&areaid=<%=areaid%>&schooltypeid=<%=schooltypeid%>&keyword=<%=keyword%>">下一页</a>
              | <a href="find_class3.asp?page=<%=pagecount%>&areaid=<%=areaid%>&schooltypeid=<%=schooltypeid%>&keyword=<%=keyword%>">末页</a>
              | </td>
          </tr>
          <% if count=0 then
               response.write "</center><font color=red>对不起，没有你要找的学校！</font></cener>"
             end if
          %>
          <tr>
            <td height="30" bgcolor="#F6F6F6">&nbsp;&nbsp;你要加入的班级是哪一级的？（以入学年份为准）：
              <select name=enterdate>
                <option value="0" selected>请选择…</option>
                <option value=1970>1970年</option>
                <option value=1971>1971年</option>
                <option value=1972>1972年</option>
                <option value=1973>1973年</option>
                <option value=1974>1974年</option>
                <option value=1975>1975年</option>
                <option value=1976>1976年</option>
                <option value=1977>1977年</option>
                <option value=1978>1978年</option>
                <option value=1979>1979年</option>
                <option value=1980>1980年</option>
                <option value=1981>1981年</option>
                <option value=1982>1982年</option>
                <option value=1983>1983年</option>
                <option value=1984>1984年</option>
                <option value=1985>1985年</option>
                <option value=1986>1986年</option>
                <option value=1987>1987年</option>
                <option value=1988>1988年</option>
                <option value=1989>1989年</option>
                <option value=1990>1990年</option>
                <option value=1991>1991年</option>
                <option value=1992>1992年</option>
                <option value=1993>1993年</option>
                <option value=1994>1994年</option>
                <option value=1995>1995年</option>
                <option value=1996>1996年</option>
                <option value=1997>1997年</option>
                <option value=1998>1998年</option>
                <option value=1999>1999年</option>
                <option value=2000>2000年</option>
                <option value=2001>2001年</option>
                <option value=2000>2002年</option>
                <option value=2001>2003年</option>
              </select>
              <input type="submit" name="Submit" value="加入">
            </td>
          </tr>
          <tr>
            <td height="28" bgcolor="#FFFFFF" align="center"> <br>
              <table width="560" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td class="topic"><font color="#993300">注意：<br>
                    　　如果你发现你的学校名称注册错误后学校重复注册，请发信通知管理员。为避免重复注册，给同校校友造成的不必要麻烦，请注册存在的学校，如果校名有误可以告知管理员修改。
                    </font></td>
                </tr>
              </table>
              <br>
            </td>
          </tr>
        </form>
      </table>
      <br>
      <table width="600" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="20"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
          <td width="480" valign="middle" class="topic" height="30"> <font color="#CC0000">如果上表中没有你要加入的学校，请填写下面的信息创建：</font></td>
        </tr>
      </table>
      <table width="600" border="0" cellspacing="1" cellpadding="1">
        <form name="form2" method="post" action="register_school.asp">
          <tr>
            <td height="28">学校名称：
              <input type="text" name="schoolname" size="40">
            </td>
          </tr>
          <tr>
            <td height="28">学校网址： http://
              <input type="text" name="schoolweb" size="40" value="无">
            </td>
          </tr>
          <tr>   <input type="Hidden" name="provinceid" value="<%=provinceid%>">
            <td height="28">学校地址：
              <input type="text" name="schooladdr" size="40" value="不知道">
              &nbsp;&nbsp;&nbsp; 邮编：
              <input type="text" name="schoolzip" size="6" value="无">
            </td>
          </tr>
          <tr>
            <td height="28">你要加入的班级是哪一级的？（以入学年份为准）：
              <select name=enterdate>
                <option value="0" selected>请选择…</option>
                <option value=1970>1970年</option>
                <option value=1971>1971年</option>
                <option value=1972>1972年</option>
                <option value=1973>1973年</option>
                <option value=1974>1974年</option>
                <option value=1975>1975年</option>
                <option value=1976>1976年</option>
                <option value=1977>1977年</option>
                <option value=1978>1978年</option>
                <option value=1979>1979年</option>
                <option value=1980>1980年</option>
                <option value=1981>1981年</option>
                <option value=1982>1982年</option>
                <option value=1983>1983年</option>
                <option value=1984>1984年</option>
                <option value=1985>1985年</option>
                <option value=1986>1986年</option>
                <option value=1987>1987年</option>
                <option value=1988>1988年</option>
                <option value=1989>1989年</option>
                <option value=1990>1990年</option>
                <option value=1991>1991年</option>
                <option value=1992>1992年</option>
                <option value=1993>1993年</option>
                <option value=1994>1994年</option>
                <option value=1995>1995年</option>
                <option value=1996>1996年</option>
                <option value=1997>1997年</option>
                <option value=1998>1998年</option>
                <option value=1999>1999年</option>
                <option value=2000>2000年</option>
                <option value=2001>2001年</option>
                <option value=2002>2002年</option>
                <option value=2003>2003年</option>
              </select>
            </td>
          </tr>
          <tr>
            <td height="28">班级全称：
              <input type="text" name="classname" size="40">
              （班级全称必须以'班'字结尾）</td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" align="center"> <br>
              <table width="560" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td class="topic"><font color="#993300">注意： <br>
                    1.校名请写全称，以避免一个学校多个名称的情况，如用“西安市第八十五中学”，<br>
                    　而不用“西安八十五中”； <br>
                    2.用中文表达，请勿用英文； <br>
                    3.数字请用中文表示，如用“西安市第八十五中学”，而不用“西安市第85中学”。 </font></td>
                </tr>
              </table>
              <br>
            </td>
          </tr>
          <tr>
            <td height="30" align="center">
              <input type="hidden" name="schooltypeid" value="<%=schooltypeid%>">
              <input type="hidden" name="areaid" value="<%=areaid%>">
              <input type="submit" name="Submit3" value="创建">
              <input type="submit" name="Submit32" value="重写">
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
<br>
    <CENTER>
<p></p>
<!--#include file=end.htm-->  </body>
</html>