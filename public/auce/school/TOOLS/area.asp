<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
provinceid = Trim(Request.form("provinceid"))
areaname=trim(request.form("areaname"))
set rs = createobject("ADODB.recordset")
sql="select * from areainfo where provinceid='"&provinceid&"'order by seed desc"
rs.open SQL,schooldb
if rs.eof then
   areaid=provinceid&"01"
   seed=1
else
   seed=rs("seed")+1
end if
rs.close
if seed<10 then
   areaid=cstr(provinceid&"0"&seed)
else
   areaid=cstr(provinceid&seed)
end if   
sql="insert into areainfo (areaid,provinceid,areaname,seed) values ('"&areaid&"','"&provinceid&"','"&areaname&"',"&seed&")"

rs.open SQL,schooldb,1,3 
response.write "Ìí¼Ó³É¹¦"  
%>       