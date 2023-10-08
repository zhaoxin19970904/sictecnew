<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
provname = Trim(Request.form("provname"))
set rs = createobject("ADODB.recordset")
sql="select * from province order by seed desc"
rs.open SQL,schooldb
if rs.eof then
   provinceid="01"
   seed=1
else
   seed=rs("seed")+1
end if
rs.close
if seed<10 then
   provinceid=cstr("0"&seed)
else
   provinceid=cstr(seed)
end if            

sql="insert into province (provinceid,provincename,seed) values ('"&provinceid&"','"&provname&"',"&seed&")"

rs.open SQL,schooldb,1,3 
response.write "Ìí¼Ó³É¹¦"
%>