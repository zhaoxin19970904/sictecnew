<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%userid=session("userid")
schoolid=trim(request("schoolid"))
enterdate=trim(request("enterdate"))
if schoolid="" then
  response.redirect "error.asp?info=请选择学校！"
end if
if enterdate="0" then
  response.redirect "error.asp?info=请选择入校年份！"
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
<div align="center">
  <center>
<table width="760" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr> 
    <td align="center" valign="top" bgcolor="#F6F6F6"> <br>
      <br>
      <br>
      <table width="600" border="0" cellspacing="0" cellpadding="0">
        <tr align="left" valign="top"> 
          <td height="26" width="20" valign="middle"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
          <td height="20" valign="middle">请选择加入下面的班级加入：</td>
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
      <table width="600" border="0" cellspacing="1" cellpadding="1">
        <form name="form1" method="post" action="find_class5.asp">
          <%
                    PageSize=5
                    set rs=Server.CreateObject("ADODB.RecordSet")
                    sql="select count(*) as a from classinfo where classid like '"&schoolid&"%' and enterdate='"&enterdate&"'"  
                    rs.open SQL,schooldb,1,3
                    count=rs("a")
                    PageCount=CInt(rs("a")/PageSize+0.5)
	            rs.Close
	            if page=1 then
	       SQL="select top 5 * from classinfo where classid like '"&schoolid&"%' and enterdate='"&enterdate&"' order by classid  "
 	    else
 	       SQL="select top 5 * from classinfo where classid like '"&schoolid&"%' and enterdate='"&enterdate&"' and classid not in (Select top "&Cstr(PageSize*(page-1))&" classid from classinfo where classid like '"&schoolid&"%' and enterdate='"&enterdate&"' order by classid ) order by classid" 
            end if
            
            rs.open SQL,schooldb
            while not rs.eof 
            
          %>
          <tr> 
            <td height="28" bgcolor="#EBEBEB"> 
              <input type="radio" name="classid" value="<%=rs("classid")%>">
              <%=rs("classname")%></td>
          </tr>
          <%
            rs.movenext
            wend
            rs.close
          %>
          <tr> 
            <td height="28" bgcolor="#CCCCCC" align="center" class="topic">当前页数<%=page%>/<%=pagecount%>页 
              | <a href="find_class4.asp?page=1&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">第一页</a> 
              | <a href="find_class4.asp?page=<%=page-1%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">上一页</a> 
              | <a href="find_class4.asp?page=<%=page+1%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">下一页</a> 
              | <a href="find_class4.asp?page=<%=pagecount%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">末页</a> 
              | </td>
          </tr>
          <% if count=0 then
               response.write "</center><font color=red>对不起，没有你要找的班级！</font></cener>"
             end if
          %>
          <tr> 
            <td height="30" bgcolor="#EBEBEB" align="center">&nbsp;&nbsp;请选择身份加入： 
              <select name="userstatus">
                <option selected value="成员">成员</option>
                <option value="教师">教师</option>
                <option value="好友">好友</option>
              </select>
              <input type="submit" name="Submit" value="加入">
            </td>
          </tr>
        </form>
      </table>
      <br>
      <table width="600" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="20"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
          <td width="480" valign="middle" class="topic" height="30"> <font color="#CC0000">如果上表中没有你要加入的班级，请填写下面的信息创建：</font></td>
        </tr>
      </table>
      <table width="600" border="0" cellspacing="1" cellpadding="1">
        <form name="form2" method="post" action="register_class.asp">
          <tr> 
            <td height="28">班级全称： 
              <input type="text" name="classname" size="40">
              （班级全称必须以'班'字结尾）</td>
          </tr>
          <tr> 
            <td height="30" align="center"> 
              <input type="hidden" name="enterdate" value="<%=enterdate%>">
              <input type="hidden" name="schoolid" value="<%=schoolid%>">
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
  </center>
</div>
<br>
<br>
     <CENTER>
<p></p>
<!--#include file=end.htm-->    </body>
</html>