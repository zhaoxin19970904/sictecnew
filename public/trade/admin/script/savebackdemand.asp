<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Save Back Inquiry</title>
</head>
<body>
<%
set rs=server.CreateObject("adodb.recordset")
id=request.QueryString("id")
ask_subject=request.Form("ask_subject")
ask_date=request.Form("ask_date")
ask_product=request.Form("ask_product")
ask_toknow=request.Form("ask_toknow")
back_subject=request.Form("back_subject")
back_content=request.Form("back_content")
member=request.Form("member")
      sql="select memberid,feedback from demand where id="&id&""
	  rs.open sql,conn,1,3
	  member=rs("memberid")
	  rs("feedback")=true
	  rs.update
	  rs.close
	  
sql="select * from back"
rs.open sql,conn,1,3
rs.addnew
rs("ask_member")=member
rs("ask_subject")=ask_subject
rs("ask_date")=ask_date
rs("ask_product")=ask_product
rs("ask_toknow")=ask_toknow
rs("back_subject")=back_subject
rs("back_content")=back_content
rs("back_date")=now()
rs("back_admin")=session("admin_name")
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script language="JavaScript">
alert("回复完成!")
location.href="viewdemand.asp"
</script>

</body>
</html>
