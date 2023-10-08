<!-- #Include File=check.asp -->
<%
user=request("user")
if request("form1")<>"" then
set rs = Server.CreateObject("ADODB.Recordset")
sqltext = "select mess From user where user='"&request.form("user")&"' "
rs.open sqltext,conn,1,1
if rs.recordcount=0 then
response.write "收信人不存在!<br>"
response.write"<input type=""button"" name=""Submit"" value=""关闭窗口"" onClick=""javascript:window.close()"">"
response.end
end if
if rs("mess")="1" then
response.write "sorry!收信人拒绝接收短信息!<br>"
response.write"<input type=""button"" name=""Submit"" value=""关闭窗口"" onClick=""javascript:window.close()"">"
response.end
end if
set rs = Server.CreateObject("ADODB.Recordset")
sqltext = "select *From mess "
rs.open sqltext,conn,3,3
from_address=address(Request.ServerVariables("REMOTE_ADDR"))
memo=request.form("ly")
memo=ubbcode(memo,vbCrlf,"<br>")
rs.addnew

rs("me")=request.cookies("renwen")("user")
rs("fd")=request.form("user")
rs("ly")=memo
rs("name")=request.form("name")
rs("date") = date
rs("time") = time
rs("ip")=from_address
rs.Update
Rs.Close
Conn.Close
Set Rs=Nothing
Set Conn=Nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('短信发送成功！谢谢你使用本程序');" & "window.close();" & "</script>"
end if
%>
<html>
<head>
<title>给<%=fd %>朋友发短信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--

function form1_onsubmit() 
{
  if(document.form1.name.value.length<1)
 {
   alert("你忘了写短信标题了");
   document.form1.name.focus();
   return false;
 }
   if(document.form1.ly.value.length<1)
 {
   alert("请填写好你要发给好友的短信");
   document.form1.ly.focus();
   return false;
  }

}

//-->
</SCRIPT>
<link rel="stylesheet" href="../images/style.css" type="text/css">
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#E0E4FE" background="backimg/03.gif">
<FORM language=javascript name=form1   onsubmit="return form1_onsubmit()"
      action=sendmess.asp method=post>
  <div align="center"></div>
  <table width="418" align=center cellpadding=3 cellspacing=1 bgcolor="#003366" class=tableborder1>
    <tbody>
      <tr align="center" bgcolor="#FFFFFF"> 
        <td height="27" colspan=3 background="backimg/bg1.gif"> <font color="#FFFFFF"> 
          给</font><font size="3" color="#FFFFFF"><b><font size="4"><font size="4"><%=user%></font></font></b></font><font color="#FFFFFF">发送短消息（请输入完整信息） </font></td>
      <tr> </tr>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1><font color="#000000"><strong>收信人</strong></font></td>
        <td valign=center class=tablebody1><font color="#FFFFFF">
          <input  name="user" value="<%=request("user")%>" size="50">
          </font><font color="#0000FF">&nbsp; </font> </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1><b>标题：</b></td>
        <td valign=center class=tablebody1><font color="#0000FF">
          <input type="text" name="name" size="50" >
          </font></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1><b>内容：</b></td>
        <td valign=center class=tablebody1><font color="#0000FF"> 
          <textarea rows="8" name="ly" cols="50" ></textarea>
          </font> </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td colspan=2 class=tablebody1> 标题最多<b>50</b>个字符，内容最多<b>300</b>个字符 </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td colspan=2 align=middle valign=center class=tablebody2> <div align="center"> 
            <input type="submit" name="form1" value="送出">
            &nbsp; 
            <input type="reset" name="Submit3" value="重写">
            &nbsp; 
            <input type="button" name="Submit" value="关闭窗口"onClick="javascript:window.close()">
          </div></td>
      </tr>
    </tbody>
  </table>
</FORM>


  