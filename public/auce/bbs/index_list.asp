<%
startime=timer()
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("bbssdata.mdb")
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connstr
  %>
  
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<%
    sqlss="select top 10 * from borecorder order by rely desc "
  set rsss=conn.execute(sqlss)
  if not rsss.eof then
  do while not rsss.eof

  %>
<table width="180" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="38" height="19"> 
      <div align="left"><img src="../Images/news.jpg" width="36" height="11"></div></td>
    <td height="19"><a href="type.asp?id=<%= rsss("id")%>" target="_blank"><%= rsss("name")%></a></td>
    <td width="20" height="19"><%= rsss("rely")%></td>
  </tr>
</table>
<%
  rsss.movenext
  loop
  set rsss=nothing
  conn.close
  end if  
  %>
</body>
</html>
