<%
	dim conn
	dim connstr
	dim db
	dim startime
	db="../../../../databases/mail.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
	'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(db)
	conn.Open connstr
    startime=timer()
%>