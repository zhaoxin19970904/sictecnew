<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Product Detail</title>
<style type=text/css>
.text{
   font-size:12px;
}
</style>
</head>
<body>
<%
i=1
id=request.QueryString("id")
sql="select * from products where id="&id&""
set rs=conn.execute(sql)
modelno=rs("modelno")
p_type=ltrim(rtrim(rs("type")))
intro=rs("introduction")
photo=rs("photo")
rs.close
sql="select type_code,type from type"
set rs=conn.execute(sql)
do while not rs.eof 
   redim preserve type_code(i),pv_type(i)
    type_code(i)=rs("type_code")
	pv_type(i)=rs("type")
  if not rs.eof then rs.movenext
  i=i+1
loop
rs.close
pvs=ubound(type_code)
for i=1 to pvs
   if p_type=type_code(i) then
      pro_type=pv_type(i)
	  exit for
	end if
next
%>


<table width="363" height="211" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td background="images/line_spacer.gif" height="1"></td>
</tr>
  <tr>
    <td width="376"><table width="361" height="95" border="0" cellpadding="0" cellspacing="0" class="text">
        <tr> 
          <td width="185" rowspan="2">
		  <div align="center">
		  <a href="proimages/<%=photo%>" target="_blank">
		  <img src="proimages/<%=photo%>" width="120" height="104" border=0>
		  </A>
		  </div></td>
          <td width="176"><font color="#6699cc">Type:</font> <%=pro_type%></td>
        </tr>
        <tr> 
          <td><font color="#6699cc">Model:</font> <%=modelno%></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="332" border="0" align="center" cellpadding="0" cellspacing="0" class="text">
        <tr>
          <td width="332" height="16"><div align="center"><font color="#6699cc">Product 
              Introduction:</font></div></td>
        </tr>
        <tr>
          <td height="26" valign="bottom">
            <div align="center">&nbsp;&nbsp;&nbsp;&nbsp;<%=intro%></div></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
conn.close
set conn=nothing
%>
</body>
</html>
