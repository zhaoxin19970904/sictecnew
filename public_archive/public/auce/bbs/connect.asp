<%@ language="vbscript"%>
<%
option explicit
'on error resume next
dim conn,connstr,startime
startime=timer()
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("bbssdata.mdb")
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connstr
%>