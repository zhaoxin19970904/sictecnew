
<%
if session("admin_name")="" then response.end
%>
<!--#include file="conn.asp"-->
<!--#include file="check.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
.unnamed1 {
	FONT-SIZE: 9pt
}
</STYLE>
</head>
<%
    sql="select * from title"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,3
	
%>
<body topmargin="100">
<div align="center"><font color="#000066" size="5">产品总类管理 <br>
  <img src="../../images/line_01.jpg" width="338" height="20"> </font></div>
<form method="POST" action="dealclass.asp" onsubmit="return del()">
<input type="hidden" name="options" value>
  <table border="0" cellpadding="0" cellspacing="0" width="44%" align="center"  class="unnamed1">

    <tr>
      <td>
            <p align="center">类别： 
              <select name="subject" size="1" style="font-size: 9pt">
                <%
	if rs.eof or rs.bof then
		 response.write "<option value='" +  "请增加类别" + "'>" + "请增加类别" + "</option>"
	  else
         Do while not rs.eof
             response.write "<option value='" +  trim(rs("title_id")) + "'>" + rs("title_type") + "</option>"
             rs.MoveNext
         Loop
	end if
%>
              </select>
      </td>
    </tr>
    <tr align="center">
      <td>
      </td>
    </tr>
    <tr align="center">
       <td height="43"> 
       <p>重命名为： 
          <input type="text" name="reTitle" size="20">
      <input type="submit" value="改名" name="B1" onclick="form.options.value='rename'">
      </td>
    </tr>
    <tr align="center">
      <td>
      </td>
    </tr>
    <tr align="center">
      <td><p>&nbsp;</td>
    </tr>
  </table>
</form>
<%
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
%>