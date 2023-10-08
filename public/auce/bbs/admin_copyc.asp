<!--#include file="admin_conn.asp" -->
<%
dim lttitle
if request.Form("Submit")<>"" then
call addpagehtm()
end if
sub addpagehtm()
dim rs,sqltext
if request.Form("tophtm")="" or request.Form("copyc")="" or request.Form("ltname")="" or request.Form("offtitle")="" or request.Form("badwords")="" then
lttitle="<li><b>!操作不成功</b><li>注意设定每项为必填项.如不要有显示只要加空格的HTML代码!"
exit sub
end if
set rs=server.createobject("adodb.recordset")
sqltext="select * from admin_copyc where id="&request.Form("id")
rs.open sqltext,conn,1,3
rs("badwords")=request.Form("badwords")
rs("tophtm")=request.Form("tophtm")
rs("copyc")=request.Form("copyc")
rs("ltname")=request.Form("ltname")
rs("offtitle")=request.Form("offtitle")
rs.update
lttitle="<li><b>!操作成功</b><li>论坛基本记定数据更新成功"
end sub
dim rs,sqltext
set rs=conn.execute("select tophtm,copyc,ltname,id,offlt,offtitle,badwords from admin_copyc")
if not rs.eof then
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>网页版面设定</title>
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>

<body>
      <form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>">
        
  <table width="573" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
    <tr align="center" background="../backimg/bg1.gif"> 
      <td height="28" colspan="2"><font color="#FFFFFF"><strong>网页版面设定</strong></font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"><%=lttitle%> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="107">主论坛状态</td>
      <td> <%if rs("offlt")=0 then 
	  response.Write("<a href=admin_cz.asp?czlb=offlt&id="&rs("id")&" title=点击关闭论坛 onclick=""{if(confirm('你确实要关闭论坛吗，确定吗?')){return true;}return false;}""> <b>开放</b></a>")
	  else
	  response.Write("<a href=admin_cz.asp?czlb=offlt&id="&rs("id")&" title=点击开放论坛 onclick=""{if(confirm('你确实要开放论坛吗，确定吗?')){return true;}return false;}""><b>关闭</b>---<font color=red>警告论坛处于关闭状态</font></a>")
	  end if
	  %>
        ---点击开放,关闭论坛 </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="107">论坛名称</td>
      <td><input name="ltname" type="text" id="ltname3" value="<%=rs("ltname")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="107">论坛过虑字符</td>
      <td><textarea name="badwords" cols="50" rows="6" id="badwords"><%=rs("badwords")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="107">论坛关闭提示<br>
        (支持HTLM)</td>
      <td><textarea name="offtitle" cols="50" rows="6" id="offtitle"><%=rs("offtitle")%></textarea> 
        <br> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>论坛顶部记置<br>
        (支持HTLM) </td>
      <td width="451"><textarea name="tophtm" cols="50" rows="6"><%=rs("tophtm")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>论坛版权信息<br>
        (支持HTLM) </td>
      <td> <textarea name="copyc" cols="50" rows="6"><%=rs("copyc")%></textarea> 
        <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;</td>
      <td align="center"> <input type="submit" name="Submit" value=" 提 交 "> &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit2" value=" 重 置 "> </td>
    </tr>
  </table>
        </form>
<%
end if
%>		
</body>
</html>
