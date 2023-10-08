<%
function search()
  key=ltrim(rtrim(request("keywords")))
  search_type=request("type")
  if search_type="" then
	  sql="select * from products where modelno like '%"&key&"%' order by up_date desc"
	  search=sql 
	  exit function
  end if
      sql="select * from products where modelno like '%"&key&"%' and type='"&search_type&"' order by up_date desc"
      set rs=conn.execute(sql)
   if rs.eof then
      rs.close
	  sql="select * from products where introduction like '%"&key&"%' and type='"&search_type&"' order by up_date desc"
	  search=sql
	  exit function
	 else
	  rs.close
	  search=sql 
	  exit function
	end if
end function
%>

