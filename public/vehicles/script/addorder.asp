<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Add Order</title>
</head>
<body>
<%
  order_type=request.form("type")
  model=request.form("model")
  quantity=request.form("quantity")
  delivery_time=request.form("delivery_time")
  contact_person=request.form("contact_person")
  tel=request.form("tel")
  fax=request.form("fax")
  address=request.form("address")
  email=request.form("email")
  notes=request.form("notes")
  set rs=server.createobject("adodb.recordset")
  sql="select * from order_"
  rs.open sql,conn,1,3
  rs.addnew
    rs("type")=order_type
	rs("model")=model
	rs("quantity")=quantity
	rs("delivery_time")=delivery_time
	rs("contact_person")=contact_person
	rs("tel")=tel
	rs("fax")=fax
	rs("address")=address
	rs("email")=email
	rs("notes")=notes
	rs("order_date")=date()
  rs.update
  rs.close
  set rs=nothing
  conn.close
%>
<script language="javascript">
  alert("Successfully!")
  location.href="../order.asp"
</script>
</body>
</html>
