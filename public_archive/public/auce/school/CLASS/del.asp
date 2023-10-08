<%
set conn=server.CreateObject("ADODB.connection")
set   rs=server.createobject("adodb.recordset")
conn.ConnectionString="driver={Microsoft Access Driver (*.mdb)};DBQ=" &server.MapPath("pic/images.mdb") & ";uid=;PWD=;"
conn.Open
sql="delete from images where id="&request("id")
conn.execute sql
conn.close
set conn=nothing
response.write"<SCRIPT language=JavaScript>alert('É¾³ý³É¹¦£¡');history.go(-1)</script>"
%>

