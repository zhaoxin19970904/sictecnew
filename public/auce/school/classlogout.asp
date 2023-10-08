<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=session("userid")
%>
<html>
<head>

	<title>мкЁЖпёсяб╪</title>
</head>
<meta http-equiv=refresh content="0; url=http://www.88com.cn">
<body>
<br><br><br>
<table width="468" cellspacing="0" cellpadding="0" border="0" align=center>
<tr><td>
<%
set rs = Server.CreateObject("ADODB.recordset")
Session.Abandon
SQL = "select * from Online where UserID='"&UserID&"'"
rs.Open SQL,DBParams,1,3
if not rs.EOF then
   rs.Delete
end if
rs.close
 response.write("<center></center>")
%>
</td>

</table>
</body>
</html>