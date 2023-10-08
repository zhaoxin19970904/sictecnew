<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<script language="javascript">

function selectschool(curid,schoolid){


                window.location.href="delete3.asp?enterdate="+curid+"&schoolid="+schoolid
    }
function deleteclass(curid){
               if (confirm("你确定要注销这个班级吗？"))
    {

        window.location.href="deleteclass.asp?classid="+curid
    }
    }
</script>
<%
if session("pass")<>"ok" then
   response.redirect "../error.asp?info=您无权进入！"
end if
schoolid=trim(request("schoolid"))
enterdate=trim(request("enterdate"))
if enterdate="" or enterdate="0" then
   temp=""
else
   temp=" and enterdate = '"&enterdate&"'"
end if
set rs = createobject("ADODB.recordset")
set rss = createobject("ADODB.recordset")
provinceid=left(schoolid,2)
areaid=left(schoolid,4)
sql="select * from province where provinceid='"&provinceid&"'"
rs.open SQL,schooldb
if not rs.eof then
   provincename=rs("provincename")
end if
rs.close
sql="select * from areainfo where areaid='"&areaid&"'"
rs.open SQL,schooldb
if not rs.eof then
   areaname=rs("areaname")
end if
rs.close
sql="select * from schoolinfo where schoolid='"&schoolid&"'"
rs.open SQL,schooldb
if not rs.eof then
   schoolname=rs("schoolname")
end if
rs.close
%>
<html>
<head>
<title>商务校友录-班级列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
</head>
<CENTER>
<!--#include file=../top2.htm-->

<body bgcolor="#FFFFFF" text="#000000" topmargin="5">
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="1" height="300" background="../school_images/vline.gif"></td>
    <td bgcolor="#F6F6F6" align="center" valign="top" width="560">
       <%
            if Request("Page")<>"" then
               Page = CLng(Request("Page"))
            end if
            If Page < 1 Then
                Page = 1
            end if
            PageSize=20
            sql="select count(*) as a from classinfo where schoolid = '"&schoolid&"' "&temp
            rs.open SQL,schooldb,1,3
            count=rs("a")
            PageCount=CInt(rs("a")/PageSize+0.5)
        rs.Close
       %>
      <table width="760" border="0" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td height="40" valign="bottom" width="150">　</td>
          <td height="40" valign="bottom" width="360" align="left"><%=provincename%><%=areaname%><font color=red><%=schoolname%></font>注册班级列表，共有<%=count%>个班级</td>
        </tr>
      </table>
      <br>
      <table width="760" border="0" cellspacing="0" cellpadding="0" bgcolor="#226886" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
            <td class="topic" align="center" height="36" bgcolor="#CCCCCC">选择你要查看的班级的入学年份：

              <select name=enterdate onchange="javascript:selectschool(this.options[selectedIndex].value,'<%=schoolid%>');">
                <option value="0" selected>请选择…</option>
                <option value="1970">1970年</option>
                <option value="1971">1971年</option>
                <option value="1972">1972年</option>
                <option value="1973">1973年</option>
                <option value="1974">1974年</option>
                <option value="1975">1975年</option>
                <option value="1976">1976年</option>
                <option value="1977">1977年</option>
                <option value="1978">1978年</option>
                <option value="1979">1979年</option>
                <option value="1980">1980年</option>
                <option value="1981">1981年</option>
                <option value="1982">1982年</option>
                <option value="1983">1983年</option>
                <option value="1984">1984年</option>
                <option value="1985">1985年</option>
                <option value="1986">1986年</option>
                <option value="1987">1987年</option>
                <option value="1988">1988年</option>
                <option value="1989">1989年</option>
                <option value="1990">1990年</option>
                <option value="1991">1991年</option>
                <option value="1992">1992年</option>
                <option value="1993">1993年</option>
                <option value="1994">1994年</option>
                <option value="1995">1995年</option>
                <option value="1996">1996年</option>
                <option value="1997">1997年</option>
                <option value="1998">1998年</option>
                <option value="1999">1999年</option>
                <option value="2000">2000年</option>
                <option value="2001">2001年</option>
              </select>

            </td>
        </tr>
      </table>
      <br>

      <table width="760" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <form name="form1" method="post" action="delete3.asp">
        <tr>
          <td height="20" class="topic" width="80">班级名称</td>
          <td height="26" class="topic" width="430" align="right" valign="top">共<%=pagecount%>页
            | <a href="delete3.asp?page=1&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">第一页</a>
            | <a href="delete3.asp?page=<%=page-1%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">上一页</a>
            | <a href="delete3.asp?page=<%=page+1%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">下一页</a>
            | <a href="delete3.asp?page=<%=pagecount%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">末页</a> | 到
            <input type="hidden" name="schoolid" value=<%=schoolid%>>
            <input type="hidden" name="enterdate" value=<%=enterdate%>>
            <input type="text" name="page" size="2">
            页
            <input type="submit" name="Submit2" value="GO">
          </td>
        </tr>
        <tr>
          <td bgcolor="#B4C7D4" height="1" colspan="2"></td>
        </tr>
      </form>
      </table>
      <br>
      <table width="760" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td class="topic" valign="top"> <%
            if page=1 then
           SQL="select top 20 * from classinfo where schoolid = '"&schoolid&"' "&temp&" order by classid  "
         else
            SQL="select top 20 * from  classinfo where schoolid = '"&schoolid&"' "&temp&" and classid not in (Select top "&Cstr(PageSize*(page-1))&" classid from classinfo where schoolid = '"&schoolid&"' "&temp&" order by classid  ) order by classid"
            end if

            rs.open SQL,schooldb
            while not rs.eof
            sql1="select count(*) as membercount from userjoinclassinfo where classid='"&rs("classid")&"'"
            rss.open SQL1,schooldb,1,3
            membercount=rss("membercount")
            rss.close
          %> <a href="../class/class_index1.asp?classid=<%=rs("classid")%>" target="_blank"><%=rs("classname")%>(已有<%=membercount%>名成员)</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:deleteclass('<%=rs("classid")%>')">注销该班级</a><br>
           <%rs.movenext
           wend
           rs.close
         %>
          </td>
        </tr>
      </table>
      <br>
      <table width="760" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
         <form name="form2" method="post" action="delete3.asp">
        <tr>
          <td height="26" class="topic" width="430" align="right" valign="top">共<%=pagecount%>页
            | <a href="delete3.asp?page=1&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">第一页</a>
            | <a href="delete3.asp?page=<%=page-1%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">上一页</a>
            | <a href="delete3.asp?page=<%=page+1%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">下一页</a>
            | <a href="delete3.asp?page=<%=pagecount%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">末页</a> | 到
              <input type="hidden" name="schoolid" value=<%=schoolid%>>
            <input type="hidden" name="enterdate" value=<%=enterdate%>>
            <input type="text" name="page" size="2">
            页
            <input type="submit" name="Submit2" value="GO">
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
<!--#include file=../end.htm--> 
      </body>
</html>