<%
set connGraph=server.CreateObject("ADODB.connection")
connGraph.ConnectionString="driver={Microsoft Access Driver (*.mdb)};DBQ=" &server.MapPath("images.mdb") & ";uid=;PWD=;"
connGraph.Open
set rec=server.createobject("ADODB.recordset")
strsql="select img from images where id=" & trim(request("id"))
rec.open strsql,connGraph,1,1
Response.ContentType = "image/*"
Response.BinaryWrite rec("img").getChunk(7500000)
rec.close
set rec=nothing
set connGraph=nothing
%>

