<%
function search()
  key=request.Form("key") 
  s_title_type=request.Form("title_type")
  'set rs=server.CreateObject("adodb.recordset")
if s_title_type="" then
   sql="select id,type,photo,name,model,type_id from products where name like '%"&key&"%' order by up_date desc"
   set rs=conn.execute(sql)
   if rs.eof then
      rs.close
	  sql="select id,type,photo,name,model,type_id from products where key_words like '%"&key&"%' order by up_date desc"
      set rs=conn.execute(sql)
	 else
	 rs.close
	  search=sql 
	  exit function
   end if
	    if rs.eof then
           rs.close
	       sql="select id,type,photo,name,model,type_id from products where model like '%"&key&"%' order by up_date desc"
           set rs=conn.execute(sql)
		  else
		   rs.close
	       search=sql
		   exit function
	     end if
		    if rs.eof then
               rs.close
	           sql="select id,type,photo,name,model,type_id from products where key_specification like '%"&key&"%' order by up_date desc"
               set rs=conn.execute(sql)
			  else
			   rs.close
               search=sql
			   exit function
	        end if
			    if rs.eof then
                   rs.close
	               response.Redirect("result_search.asp")
                   response.end
				  else
				  rs.close
	               search=sql
				   exit function
	            end if
	  
else


         sql="select id,type,photo,name,model,type_id from products where name like '%"&key&"%' and title_type='"&s_title_type&"' order by up_date desc"
         set rs=conn.execute(sql)
   if rs.eof then
      rs.close
	  sql="select id,type,photo,name,model,type_id from products where key_words like '%"&key&"%' and title_type='"&s_title_type&"' order by up_date desc"
      set rs=conn.execute(sql)
	 else
	  rs.close
	  search=sql
	  exit function
   end if
	    if rs.eof then
           rs.close
	       sql="select id,type,photo,name,model,type_id from products where model like '%"&key&"%' and title_type='"&s_title_type&"' order by up_date desc"
           set rs=conn.execute(sql)
		  else
		   rs.close
	       search=sql
		   exit function
	     end if
		    if rs.eof then
               rs.close
	           sql="select id,type,photo,name,model,type_id from products where key_specification like '%"&key&"%' and title_type='"&s_title_type&"' order by up_date desc"
               set rs=conn.execute(sql)
			  else
			   rs.close
	           search=sql
			   exit function
	        end if
			    if rs.eof then
                   rs.close
	               response.Redirect("result_search.asp")
                   response.end
				  else
				   rs.close
	               search=sql
				   exit function
	            end if
end if

end function
%>

