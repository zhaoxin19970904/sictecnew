<!--#include file="CONN.ASP" -->
<%
if not userislogin then
errmess=errmess&error1
call founderror(errmess)
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=loadcopyc("ltname")%>--短信息</title>
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>

<body  onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()">
<form language=javascript name="messager" action=messanger.asp method=post>
  <table width="898" border="0" align=center cellpadding=3 cellspacing=1 bgcolor="#92b9fb" class=tableborder1>
    <tbody>
      <tr> 
        <td height="27" colspan=3 background="backimg/bg1.gif"> <b><font color="#FFFFFF">发送短信息 
          </font></b> </td>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1>收信人</td>
        <td valign=center class=tablebody1>
          <input name="touser" type="text" id="touser" size="50" maxlength="200">
          <SELECT name=fonta onchange="document.messager.touser.value+=fonta.value">
            <OPTION selected value="">收件人</OPTION>
<%
call listmyfriend()
sub listmyfriend()
dim rs
set rs=conn.execute("select me,fd from fd where me='"&loginuser&"'")
do while not rs.eof
response.Write("<OPTION value="""&rs("fd")&","">"&rs("fd")&"</OPTION>")
rs.movenext
loop
end sub
%>
          </SELECT> </td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td valign=top class=tablebody1>标题：</td>
        <td valign=center class=tablebody1><font color="#0000FF">
          <input name="name" type="text" size="50" maxlength="50" >
          </font></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1>内容：</td>
        <td valign=center class=tablebody1> <textarea rows="8" name="ly" cols="50" ></textarea> 
        </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td colspan=2 class=tablebody1> 说明：<br>
          ① 您可以使用Ctrl+Enter键快捷发送短信<br>
          ② 可以用英文状态下的逗号将用户名隔开实现群发，最多5个用户<br>
          ③ 标题最多50个字符，内容最多4000个字符</td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td colspan=2 align=middle valign=center class=tablebody2> <div align="center"> 
            <font color="#0000FF"> 
            <input name=FORM1 type=submit  value="发 送">
            &nbsp; 
            <input type="reset" name="Reset" value="重 填">
            &nbsp;&nbsp; 
            <input type="button" name="Submit2" value="关 闭"onClick="javascript:window.close()">
             &nbsp;&nbsp; 
            <input type="button" name="Submit" value="聊天记录">
            </font></div></td>
      </tr>
  </table>
</form>
</body>
</html>
