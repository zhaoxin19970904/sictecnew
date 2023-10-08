<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<script language="javascript">

function selecttype(curid,provinceid,areaid){


                window.location.href="local.asp?schooltypeid="+curid+"&provinceid="+provinceid+"&areaid="+areaid

    }
function selectarea(curid,provinceid,schooltypeid){


                window.location.href="local.asp?areaid="+curid+"&provinceid="+provinceid+"&schooltypeid="+schooltypeid

    }

</script>
<%
provinceid=trim(request("provinceid"))
schooltypeid=trim(request("schooltypeid"))
areaid=trim(request("areaid"))
set rs = createobject("ADODB.recordset")
set rss = createobject("ADODB.recordset")
if schooltypeid="" or schooltypeid="0" then
   if areaid="" or areaid="0" then
      temp=""
   else
      temp=" and areaid='"&areaid&"'"
   end if
else
   if areaid="" or areaid="0" then
      temp=" and schooltypeid="&clng(schooltypeid)
   else
      temp=" and areaid='"&areaid&"' and  schooltypeid="&clng(schooltypeid)
   end if
end if
if schooltypeid="" then
   schooltypeid=0
end if
if areaid="" then
   areaid=0
end if
sql="select * from province where provinceid='"&provinceid&"'"
rs.open SQL,schooldb
if not rs.eof then
   provincename=rs("provincename")
end if
rs.close
%>
<html>
<head>
<title>商务校友录-学校列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<STYLE type=text/css>
</STYLE>
<LINK href="scss.css" rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="5"><CENTER>
<!--#include file=top2.htm-->

<div align="center">
  <center>
<table width="760" border="0" cellspacing="1" bordercolor="#000000">
  <tr>
    <td bgcolor="#F6F6F6" align="center" valign="top" width="324">
      <table width="760" border="0" cellpadding="2">
        <tr>
          <td height="40" valign="bottom" width="150">　</td>
            <%
            sql="select count(*) as a from schoolinfo where schoolid like '"&provinceid&"%'"
            rs.open SQL,schooldb,1,3
            schoolcount=rs("a")
            rs.close
            %>
          <td height="40" valign="bottom" width="360" align="left">您现在浏览的是<font color="red"><%=provincename%></font>注册学校列表，共有<%=schoolcount%>所学校</td>
        </tr>
      </table>
      <br>
      <table width="760" border="0" cellspacing="1" bgcolor="#5B5B5B">
        <tr>
            <td class="topic" align="center" height="36" bgcolor="#CCCCCC">
            <font color="#000000">分类查看
              &gt;&gt; 学校类型： </font>
              <select name="schooltypeid" onchange="javascript:selecttype(this.options[selectedIndex].value,'<%=cstr(provinceid)%>','<%=areaid%>');">
              <option>请选择</option>
              <%
              sql="select * from schooltype order by schooltypeid"
              rs.open SQL,schooldb
              while not rs.eof
              selected1=""
              if cint(schooltypeid)=cint(rs("schooltypeid")) then
                 selected1="selected"
              end if
              %>
              <option <%=selected1%> value=<%=rs("schooltypeid")%>><%=rs("schooltypename")%></option>
              <%
               rs.movenext
               wend
               rs.close
              %>
            </select><font color="#000000">
            &nbsp;&nbsp;学校所在地区： </font>
            <select name="areaid" onchange="javascript:selectarea(this.options[selectedIndex].value,'<%=cstr(provinceid)%>',<%=schooltypeid%>);">
             <option>请选择</option>
             <%
             sql="select * from areainfo where provinceid='"&provinceid&"' order by seed "
             rs.open SQL,schooldb
             while not rs.eof
             selected2=""
             if areaid=rs("areaid") then
                selected2="selected"
             end if
              %>
             <option <%=selected2%> value="<%=rs("areaid")%>"><%=rs("areaname")%></option>
             <%
               rs.movenext
               wend
               rs.close
              %>
            </select><font color="#000000">
              &nbsp;&nbsp; </font>

            </td>
        </tr>
      </table>
      <br>
      <%
            if Request("Page")<>"" then
               Page = CLng(Request("Page"))
            end if
            If Page < 1 Then
                Page = 1
            end if
            PageSize=20
            sql="select count(*) as a from schoolinfo where schoolid like '"&provinceid&"%' "&temp
            rs.open SQL,schooldb,1,3
            count=rs("a")
            PageCount=CInt(rs("a")/PageSize+0.5)
        rs.Close
       %>
      <table width="760" border="0" cellspacing="0" cellpadding="0">
        <form name="form1" method="post" action="local.asp">
        <tr>
          <td height="20" class="topic" width="80">&nbsp;<font color="#000000">学校名称</font></td>
          <td height="26" class="topic" width="430" align="right" valign="top">
          <font color="#000000">共<%=pagecount%>页
            | </font> <a href="local.asp?page=1&provinceid=<%=provinceid%>&schooltypeid=<%=schooltypeid%>&areaid=<%=areaid%>">
          <font color="#000000">第一页</font></a><font color="#000000"> | </font>
             <a href="local.asp?page=<%=page-1%>&provinceid=<%=provinceid%>&schooltypeid=<%=schooltypeid%>&areaid=<%=areaid%>">
          <font color="#000000">上一页</font></a><font color="#000000"> | </font>
             <a href="local.asp?page=<%=page+1%>&provinceid=<%=provinceid%>&schooltypeid=<%=schooltypeid%>&areaid=<%=areaid%>">
          <font color="#000000">下一页</font></a><font color="#000000">
            | </font> <a href="local.asp?page=<%=pagecount%>&provinceid=<%=provinceid%>&schooltypeid=<%=schooltypeid%>&areaid=<%=areaid%>">
          <font color="#000000">末页</font></a><font color="#000000"> | 到 </font>
            <input type="text" name="page" size="2"><font color="#000000">
            页 </font>
             <input type="hidden" name="provinceid" value=<%=provinceid%>>
             <input type="hidden" name="areaid" value=<%=areaid%>>
             <input type="hidden" name="schooltypeid" value=<%=schooltypeid%>>
            <input type="submit" name="Submit2" value="GO"><font color="#000000">
          </font>
          </td>
        </tr>
        <tr>
          <td bgcolor="#B4C7D4" height="1" colspan="2"></td>
        </tr>
      </form>
      </table>
      <br>
      <table width="760" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td class="topic">
            <%
            if page=1 then
           SQL="select top 20 * from schoolinfo where schoolid like '"&provinceid&"%' "&temp&" order by schoolid  "
         else
            SQL="select top 20 * from schoolinfo where schoolid like '"&provinceid&"%' "&temp&" and schoolid not in (Select top "&Cstr(PageSize*(page-1))&" schoolid from schoolinfo where schoolid like '"&provinceid&"%' "&temp&" order by schoolid  ) order by schoolid"
            end if

            rs.open SQL,schooldb
            while not rs.eof
            sql1="select count(*) as classcount from classinfo where schoolid='"&rs("schoolid")&"'"
            rss.open SQL1,schooldb,1,3
            classcount=rss("classcount")
            rss.close
          %><font color="#000000"> </font>
            &nbsp;<a href="classlist.asp?schoolid=<%=rs("schoolid")%>"><%=rs("schoolname")%><font color="#000000">(已有<%=classcount%>个班)</font></a><font color="#000000"><br>
         <%rs.movenext
           wend
           rs.close
         %> </font>
          </td>
        </tr>
      </table>
      <br>
      <table width="760" border="0" cellspacing="0" cellpadding="0">
        <form name="form1" method="post" action="local.asp">
        <tr>
          <td bgcolor="#B4C7D4" height="1" colspan="2"></td>
        </tr><tr>
          <td height="20" class="topic" width="80"></td>
            <td height="26" class="topic" width="430" align="right" valign="bottom">
            <font color="#000000">共<%=pagecount%>页
              | </font> <a href="local.asp?page=1&provinceid=<%=provinceid%>&schooltypeid=<%=schooltypeid%>&areaid=<%=areaid%>">
            <font color="#000000">第一页</font></a><font color="#000000">
              | </font> <a href="local.asp?page=<%=page-1%>&provinceid=<%=provinceid%>&schooltypeid=<%=schooltypeid%>&areaid=<%=areaid%>">
            <font color="#000000">上一页</font></a><font color="#000000">
              | </font> <a href="local.asp?page=<%=page+1%>&provinceid=<%=provinceid%>&schooltypeid=<%=schooltypeid%>&areaid=<%=areaid%>">
            <font color="#000000">下一页</font></a><font color="#000000">
              | </font> <a href="local.asp?page=<%=pagecount%>&provinceid=<%=provinceid%>&schooltypeid=<%=schooltypeid%>&areaid=<%=areaid%>">
            <font color="#000000">末页</font></a><font color="#000000">
              | 到 </font>
              <input type="text" name="page" size="2"><font color="#000000">
            页 </font>
             <input type="hidden" name="provinceid" value=<%=provinceid%>>
             <input type="hidden" name="areaid" value=<%=areaid%>>
             <input type="hidden" name="schooltypeid" value=<%=schooltypeid%>>
            <input type="submit" name="Submit2" value="GO"><font color="#000000">
            </font>
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
  </center>
</div>
<br>
<br>
       <CENTER>
<p></p>
<!--#include file=end.htm--></body>
</html>