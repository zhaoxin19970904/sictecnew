<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css" type="text/css">
<title>Product Detail</title>
</style>
</head>
<body>
<%
i=1
id=request.QueryString("id")
sql="select * from products where id="&id&""
set rs=conn.execute(sql)
modelno=rs("modelno")
price=rs("price")
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
    <td background="images/line_spacer.gif" height="1" colspan="2"></td>
  </tr>
  <tr> 
    <td width="376" colspan="2"> <table width="361" height="223" border="0" cellpadding="0" cellspacing="0" class="text">
        <tr> 
          <td width="361"> <div align="center"> <a href="proimages/<%=photo%>" target="_blank"> 
              <img src="proimages/<%=photo%>" width="249" height="217" border=0> 
              </A> </div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="18" width="48"><font color="#6699cc" size="2">Type:</font> </td>
    <td height="18" width="315"><font color="#828653"><%=pro_type%></font>　</td>
  </tr>
  <tr> 
    <td height="19"><font color="#6699cc" size="2">Model:</font></td>
    <td height="19"><font color="#828653"><%=modelno%></font></td>
  </tr>
  <tr> 
    <td height="19" width="48"><font color="#6699cc" size="2">Price：</font></td>
    <td height="19" width="315"><font color="#828653">$&nbsp;<%=price%></font>　</td>
  </tr>
  <tr> 
    <td height="42" colspan="2"> <table width="362" border="0" align="left" cellpadding="0" cellspacing="0" class="text" height="72">
        <tr> 
          <td width="362" height="20">
<div align="left"><font color="#6699cc" size="2">Product 
              Introduction：</font></div></td>
        </tr>
        <tr> 
          <td height="52" valign="bottom"> <font color="#828653"><div align="left"><%=intro%></div></font></td>
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