<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=����_GB2312">
<title>�����Ϣ</title>
</head>
<body>
<%
 id=request.QueryString("id")
 p_type=request.form("type")
 
 apart=split(p_type,":")
 
 sql="select distinct main_id from type where type_id="&apart(0)&""
 set rs=conn.execute(sql)
 main_id=rs(0)
 rs.close
 
 sql="select distinct title_id from main where main_id="&main_id&""
 set rs=conn.execute(sql)
 title_id=rs(0)
 rs.close
 
 sql="select distinct title_type from title where title_id="&title_id&""
 set rs=conn.execute(sql)
 title_type=rs(0)
 rs.close
 
 
 
 
 p_name=request.form("name")
 model=request.form("model")
 brandname=request.form("brandname")
 region=request.form("region")
 approvals=request.form("approvals")
 Price=request.form("Price")
 main_export_markets=request.form("main_export_markets")
 Packing=request.form("Packing")
 key_specification=request.form("key_specification")
 payment_terms=request.form("payment_terms")
 lead_time=request.form("lead_time")
 cost=request.form("cost")
 manufacture=request.Form("manufacture")
 related_products=request.Form("related_products")
 key_words=request.Form("key_words")
 sample=request.Form("sample")
 Note=request.Form("Note")
 set rs=server.createobject("adodb.recordset")
 sql="select * from products where id="&id&""
 rs.open sql,conn,1,3
 'rs("type")=p_type
 
 rs("type")=apart(1)
 rs("title_type")=ltrim(rtrim(title_type))
 rs("type_id")=apart(0)
 
 rs("name")=p_name
 rs("model")=model
 rs("brandname")=brandname
 rs("region")=region
 rs("approvals")=approvals
 rs("Price")=Price
 rs("main_export_markets")=main_export_markets
 rs("Packing")=Packing
 rs("key_specification")=key_specification
 rs("payment_terms")=payment_terms
 rs("lead_time")=lead_time
 rs("cost")=cost
 rs("manufacture")=manufacture
 rs("related_products")=related_products
 rs("key_words")=key_words
 if trim(Sample)="��" then
     rs("sample")=true
   else
     rs("sample")=false
 end if
  rs("Notes")=Note
 rs.update
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 response.write "��Ʒ��Ϣ�ѳɹ��޸�<br><br>"
 response.write "<a href='viewpro.asp'>ȷ��</a>"
%>
</body>
</html>
