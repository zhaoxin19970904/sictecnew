<%
if session("user_name")<>"sictec" then
   response.redirect "index.htm"
   response.end 
end if
%>
<%
	dim conn
	dim connstr
	dim db
	db="database/sicteccity.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
'�����ķ��������ý��ϰ汾Access�����������������ӷ���
	'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(db)
	conn.Open connstr
%>

