<%
dim p_id(),modelno(),photo(),j,sql,commended_pros,c_reds
sql="select id from products where ifcommended=true"
rs.open sql,conn,1,1
c_reds=rs.recordcount
if c_reds>=4 then
    sql="select top 4 * from products where ifcommended=true"
  elseif c_reds<4 and c_reds>0 then
    sql="select * from products where ifcommended=true"
  elseif c_reds=0 then
%>
<table width="710" border="0" cellpadding="0" align="center" height="145" cellspacing="0">
  <tr> 
    <td height="145">
	 <div align="center"> <font size="2">No commended products for the moment 
        .</font> </div></td>
  </tr>
</table>
<%
elseif c_reds<>0 then
rs.close
j=1
set rs=conn.execute(sql)
do while not rs.eof
  redim preserve p_id(j),modelno(j),photo(j)
  p_id(j)=rs("id")
  modelno(j)=rs("modelno")
  photo(j)=rs("photo")
  if not rs.eof then rs.movenext
  j=j+1
loop
rs.close
commended_pros=ubound(p_id)
%>
<table width="710" border="0" cellpadding="0" align="center" height="145" cellspacing="0">
  <tr> 
    <td width="169" height="145"> 
	<%if commended_pros>=1 then%>
      <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
        <tr> 
          <td width="120">
		  <img src="proimages/<%=photo(1)%>" width="120" height="104">
		  </td>
        </tr>
        <tr> 
          <td height="25"> 
		  <div align="center">
		     <font size="2">
		       <%=modelno(1)%>
		     </font>
		  </div>
		  </td>
        </tr>
      </table>
	  <%end if%>
	  </td>
    <td width="180" height="145"> 
	<%if commended_pros>=2 then%>
      <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
        <tr> 
          <td>
		  <img src="proimages/<%=photo(2)%>" width="120" height="104">
		  </td>
        </tr>
        <tr> 
          <td height="25"> 
		  <div align="center">
		  <font size="2">
		  <%=modelno(2)%>
		  </font>
		  </div></td>
        </tr>
      </table>
	  <%end if%>
	  </td>
    <td width="180" height="145"> 
	<%if commended_pros>=3 then%>
      <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
        <tr> 
          <td>
		  <img src="proimages/<%=photo(3)%>" width="120" height="104">
		  </td>
        </tr>
        <tr> 
          <td height="25">
		   <div align="center">
		   <font size="2">
		    <%=modelno(3)%>
		   </font>
		   </div></td>
        </tr>
      </table>
	  <%end if%>
	  </td>
    <td width="181" height="145"> 
	<%if commended_pros>=4 then%>
      <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
        <tr> 
          <td>
		  <img src="proimages/<%=photo(4)%>" width="120" height="104">
		  </td>
        </tr>
        <tr> 
          <td height="25"> 
		  <div align="center">
		  <font size="2">
		  <%=modelno(4)%>
		  </font>
		  </div></td>
        </tr>
      </table>
	  <%end if%>
	  </td>
  </tr>
</table>
<%end if%>

