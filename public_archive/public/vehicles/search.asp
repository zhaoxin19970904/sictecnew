<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<!--#include file="inc/search.asp" -->
<html>
<head>
<title><!--#INCLUDE FILE="inc/title.htm" --></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
<style type="text/css">
.text{
 font-size:12px;
}
</style>
</head>
<%
  set rs=server.CreateObject("adodb.recordset")
  key1=request("keywords")
  type1=request("type")
%>

<body bgcolor="#FFFFFF" topmargin="0"  leftmargin="0"  text="#000000" background="star/dd_boot.jpg">
<table width="1000" border="0" cellspacing="0" cellpadding="0" height="509">
  <tr>
    <td valign="top" height="572" width="9"> 
      <table width="980" border="0" cellspacing="0" cellpadding="0" height="70">
        <tr> 
          <td><!--#include file="inc/top.htm" --></td>
        </tr>
      </table>
      <table width="1000" border="0" cellspacing="0" cellpadding="0" height="86" align="center">
        <tr> 
          <td background="star/dd_boot_6.jpg" valign="top"> 
            <table width="979" border="0" cellspacing="0" cellpadding="0" height="131">
              <tr> 
                <td height="114">&nbsp;</td>
              </tr>
              <tr> 
                <td height="14"> 
                  <table width="980" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="250"><img src="star/dd_boot_7.jpg" width="251" height="14"></td>
                      <td width="730"><img src="star/dd_boot_8.jpg" width="729" height="14"></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table width="996" border="0" cellspacing="0" cellpadding="0" height="540">
        <tr> 
          <td width="256" height="377" valign="top">
            <table width="256" border="0" cellspacing="0" cellpadding="0" height="251">
              <tr> 
                <td width="11" height="274">&nbsp;</td>
                <td valign="top" height="274"> 
                  <table width="256" border="0" cellspacing="0" cellpadding="0" height="251">
                    <tr> 
                      <td width="11" height="274">&nbsp;</td>
                      <td valign="top" height="274"> 
                        <table width="240" border="0" cellspacing="0" cellpadding="0" height="25">
                          <tr> 
                            <td height="60"> 
                              <div align="center"><img src="star/dd_1.jpg" width="211" height="44"></div>
                            </td>
                          </tr>
                          <tr> 
                            <td height="25"> 
                              <div align="center"> <a href="products.asp?id=1"> 
                                </a><a href="products.asp?id=1"><img src="star/dd_2_a_1.jpg" width="196" height="25" border="0"></a> 
                              </div>
                            </td>
                          </tr>
                          <tr> 
                            <td height="25"> 
                              <div align="center"> <a href="products.asp?id=2"> 
                                </a><a href="products.asp?id=2"><img src="star/dd_2_a_2.jpg" width="196" height="25" border="0"></a> 
                              </div>
                            </td>
                          </tr>
                          <tr> 
                            <td height="25"> 
                              <div align="center"> <a href="products.asp?id=3"> 
                                </a><a href="products.asp?id=3"><img src="star/dd_2_a_3.jpg" width="196" height="25" border="0"></a> 
                              </div>
                            </td>
                          </tr>
                          <tr> 
                            <td height="25"> 
                              <div align="center"> <a href="products.asp?id=4"> 
                                </a><a href="products.asp?id=4"><img src="star/dd_2_a_4.jpg" width="196" height="25" border="0"></a> 
                              </div>
                            </td>
                          </tr>
                          <tr> 
                            <td height="25"> 
                              <div align="center"> <a href="products.asp?id=5"> 
                                </a><a href="products.asp?id=5"><img src="star/dd_2_a_5.jpg" width="196" height="25" border="0"></a> 
                              </div>
                            </td>
                          </tr>
                          <tr> 
                            <td height="25"> 
                              <div align="center"> <a href="products.asp?id=6"> 
                                </a><a href="products.asp?id=6"><img src="star/dd_2_a_6.jpg" width="196" height="25" border="0"></a> 
                              </div>
                            </td>
                          </tr>
                          <tr> 
                            <td height="25"> 
                              <div align="center"> <a href="products.asp?id=7"> 
                                </a><a href="products.asp?id=7"><img src="star/dd_2_a_7.jpg" width="196" height="25" border="0"></a> 
                              </div>
                            </td>
                          </tr>
                          <tr> 
                            <td height="25"> 
                              <div align="center"> <a href="new_products.asp"> 
                                </a><a href="new_products.asp"><img src="star/dd_2_a_8.jpg" width="196" height="25" border="0"></a> 
                              </div>
                            </td>
                          </tr>
                          <tr> 
                            <td height="25"> 
                              <div align="center"><img src="star/dd_2_a_9.jpg" width="196" height="31" border="0"></div>
                            </td>
                          </tr>
                        </table>
                      </td>
                      <td width="5" height="274">&nbsp;</td>
                    </tr>
                  </table>
                </td>
                <td width="5" height="274">&nbsp;</td>
              </tr>
            </table>
          </td>
          <td height="377" valign="top">
            <table width="732" border="0" cellspacing="0" cellpadding="0" height="563">
              <tr> 
                <td width="723" height="563" valign="top">
<table width="721" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><table width="710" border="0" cellspacing="1" cellpadding="0" align="center" height="50" bgcolor="#CCCCCC">
                          <tr> 
                            <td height="50" bgcolor="#F6F6F6"> 
							<form name="form1" method="post" action="search.asp" onsubmit="return check1(form1)">
                                <table width="707" height="36" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="125" valign="bottom"> <div align="center"><strong><font color="#6699CC" size="4">Search</font></strong></div></td>
                                    <td width="165" valign="bottom"> 
                                      <%
								dim id(),v_type(),type_code(),types,i
								i=1
							    sql="select * from type"
								set rs=conn.execute(sql)
								do while not rs.eof
								 redim preserve id(i),v_type(i),type_code(i)
								 id(i)=rs("id")
								 v_type(i)=rs("type")
								 type_code(i)=rs("type_code")
								 if not rs.eof then rs.movenext
								 i=i+1
								loop
								rs.close
								types=ubound(id)
							  %>
                                      <select name="type" id="type">
									  <option value="">===All===</option>
                                        <%for i=1 to types%>
                                        <option value="<%=type_code(i)%>" <%if type_code(i)=type1 then response.Write "selected" else response.Write "" end if%>><%=v_type(i)%></option>
                                        <%
								 next
								%>
                                      </select> </td>
                                    <td width="269" valign="bottom"> <div align="center"><strong><font color="#6699CC">Keywords:</font></strong> 
                                        <input name="keywords" type="text" id="keywords" size="15" value="<%=key1%>">
                                      </div></td>
                                    <td width="148" valign="bottom"> 
									<input type="submit" name="Submit" value="Search"> 
                                    </td>
                                  </tr>
                                </table>
                              </form></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr>
                      <td height=11></td>
                    </tr>
                    <tr>
                      <td height="355" valign="top"> 
                        <%
		'sql="select * from products order by up_date desc"
		 sql=search()
		rs.open sql,conn,1,1
		pv_reds=rs.recordcount
if pv_reds=0 then
        rs.close
        response.write "<div align=center><font size='2'>Sorry,search result is empty for the moment.</font></div>"
elseif pv_reds<>0 or pv_reds<>-1 then

 inputpage=request.form("inputpage")
if len(inputpage)<>0 then
      iPage=cint(inputpage)
      rs.absolutepage=inputpage
   else   
     if len(request("page"))=0 then
         iPage=1 
       else
         iPage=request("page")
     end if
end if 

  rs.pagesize=12
  rs.absolutepage=iPage
  rs.cachesize=rs.pagesize
  i=0
dim pv_id(),pv_photo(),pv_type(),pv_modelno(),intro()
do while not rs.eof and (i<rs.pagesize)
          i=i+1
      redim preserve pv_id(i),pv_photo(i),pv_type(i),pv_modelno(i),intro(i)
          pv_id(i)=rs("id")
		  pv_photo(i)=rs("photo")
          pv_type(i)=rs("type")
		  pv_modelno(i)=rs("modelno")
		  intro(i)=rs("introduction")
		  rs.movenext
loop
	 i=1 
	 pv_ids=ubound(pv_id)
Do while i<=pv_ids
%>
                        <table width="710" border="0" cellpadding="0" align="center" height="88" cellspacing="0">
                          <tr> 
                            <td height="113"> <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
                                <tr> 
                                  <td> <a href="#" onClick="javascript:window.open('detail.asp?id=<%=pv_id(i)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                    <img src="proimages/<%=pv_photo(i)%>" width="120" height="104" border="0"> 
                                    </a> </td>
                                </tr>
                                <tr> 
                                  <td height="25"> <div align="center"> <font size="2"> 
                                      <a href="#" onClick="javascript:window.open('detail.asp?id=<%=pv_id(i)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                      <%=pv_modelno(i)%> </a> </font></div></td>
                                </tr>
                              </table></td>
                            <td height="113" width="176"> 
                              <%
			i=i+1
			if i<=pv_ids then
	      %>
                              <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
                                <tr> 
                                  <td> <a href="#" onClick="javascript:window.open('detail.asp?id=<%=pv_id(i)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                    <img src="proimages/<%=pv_photo(i)%>" width="120" height="104" border="0"> 
                                    </a> </td>
                                </tr>
                                <tr> 
                                  <td height="25"> <div align="center"><font size="2"> 
                                      <a href="#" onClick="javascript:window.open('detail.asp?id=<%=pv_id(i)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                      <%=pv_modelno(i)%> </a> </font></div></td>
                                </tr>
                              </table>
                              <%
			end if
			i=i+1
			%>
                            </td>
                            <td height="113" width="176"> 
                              <%				  
             if i<=pv_ids then
			%>
                              <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
                                <tr> 
                                  <td> <a href="#" onClick="javascript:window.open('detail.asp?id=<%=pv_id(i)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                    <img src="proimages/<%=pv_photo(i)%>" width="120" height="104" border="0"> 
                                    </a> </td>
                                </tr>
                                <tr> 
                                  <td height="25"> <div align="center"><font size="2"> 
                                      <a href="#" onClick="javascript:window.open('detail.asp?id=<%=pv_id(i)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                      <%=pv_modelno(i)%> </a> </font></div></td>
                                </tr>
                              </table>
                              <%
			end if
			i=i+1
			%>
                            </td>
                            <td height="113" width="176"> 
                              <%if i<=pv_ids then%>
                              <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
                                <tr> 
                                  <td> <a href="#" onClick="javascript:window.open('detail.asp?id=<%=pv_id(i)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                    <img src="proimages/<%=pv_photo(i)%>" width="120" height="104" border="0"> 
                                    </a> </td>
                                </tr>
                                <tr> 
                                  <td height="25"> <div align="center"> <font size="2"> 
                                      <a href="#" onClick="javascript:window.open('detail.asp?id=<%=pv_id(i)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                      <%=pv_modelno(i)%> </a> </font></div></td>
                                </tr>
                              </table>
                              <%
            end if
            i=i+1
    %>
                            </td>
                          </tr>
                        </table> 
                        <% loop %>
                      </td>
                    </tr>
                    <tr>
                      <td><table width="704" height="39" border="0" align="center" cellpadding="0" cellspacing="0" class="text">
                          <tr> 
                            <td width="704" height="39"> <form name="form2" method="post" action="" style="height:20">
                                <div align="right" class="dateformat">Page No.<font color="#6699cc"> 
                                  <%if len(inputpage)=0 and isnumeric(inputpage) then%>
                                  <%=iPage%> 
                                  <%else%>
                                  <%=inputpage%> 
                                  <%end if%>
                                  </font> &nbsp;&nbsp;&nbsp;&nbsp;Total <font color="#6699cc"><%=rs.recordcount%></font> 
                                  Products&nbsp;&nbsp;
                                  <%
							pages=rs.pagecount
							%>
                                  &nbsp;&nbsp;&nbsp;&nbsp;GO 
                                  <select name="inputpage" onChange="form2submit()">
								  <option value="" selected>select</option>
                                    <%for m=1 to pages%>
                                    <option value="<%=m%>">page:<%=m%></option>
                                    <%next%>
                                  </select>
                                  Page &nbsp;&nbsp;&nbsp;&nbsp; 
                                  <%
if cint(iPage)=1 or cint(inputpage)=1 then
%>
                                  <font class="text">First Back </font> 
                                  <%
else
%>
                                  <font class="text"><a href="search.asp?page=1&keywords=<%=key1%>&type=<%=type1%>">First 
                                  </a></font> <font class="text"><a href="search.asp?page=<%=iPage-1%>&keywords=<%=key1%>&type=<%=type1%>">Back</a></font> 
                                  <%end if%>
                                  <%
 if cint(iPage)=cint(pages) or cint(inputpage)=cint(pages) then
%>
                                  <font class="text"> 
                                  <input type="hidden" name="keywords" value="<%=key1%>">
                                  <input type="hidden" name="type" value="<%=type1%>">
                                  Next Last</font> 
                                  <%else%>
                                  <a href="search.asp?page=<%=iPage+1%>"> Next</a><a href="search.asp?page=<%=pages%>&keywords=<%=key1%>&type=<%=type1%>"> 
                                  Last</a> 
                                  <%
end if
%>&nbsp;&nbsp;&nbsp;
                                </div>
                              </form></td>
                          </tr>
                        </table>
                        <%
rs.close
end if
%>
                      </td>
                    </tr>
                    <tr>
                      <td><div align="center"><img src="star/dd_boot_9.jpg" width="710" height="27"></div></td>
                    </tr>
                    <tr>
                      <td>
<%
dim p_id(),modelno(),photo(),j,introduction(),sql,commended_pros,c_reds
sql="select id from products where ifcommended=true"
rs.open sql,conn,1,1
c_reds=rs.recordcount
if c_reds>=4 then
    sql="select top 4 * from products where ifcommended=true"
  elseif c_reds<4 and c_reds>0 then
    sql="select * from products where ifcommended=true"
end if
rs.close
j=1
set rs=conn.execute(sql)
do while not rs.eof
  redim preserve p_id(j),modelno(j),photo(j),introduction(j)
  p_id(j)=rs("id")
  modelno(j)=rs("modelno")
  photo(j)=rs("photo")
  introduction(j)=rs("introduction")
  if not rs.eof then rs.movenext
  j=j+1
loop
rs.close
commended_pros=ubound(p_id)
%>
                        <table width="710" border="0" cellpadding="0" align="center" height="130" cellspacing="0">
                          <tr> 
                            <td width="169" height="130"> 
                              <%if commended_pros>=1 then%>
                              <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
                                <tr> 
                                  <td width="120"> <a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(1)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                    <img src="proimages/<%=photo(1)%>" width="120" height="104" border="0"> 
                                    </a> </td>
                                </tr>
                                <tr> 
                                  <td height="25"> <div align="center"> <font size="2"> 
                                      <a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(1)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                      <%=modelno(1)%> </a> </font> </div></td>
                                </tr>
                              </table>
                              <%end if%>
                            </td>
                            <td width="180" height="181"> 
                              <%if commended_pros>=2 then%>
                              <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
                                <tr> 
                                  <td> <a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(2)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                    <img src="proimages/<%=photo(2)%>" width="120" height="104" border="0"> 
                                    </a> </td>
                                </tr>
                                <tr> 
                                  <td height="25"> <div align="center"> <font size="2"> 
                                      <a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(2)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                      <%=modelno(2)%> </a> </font> </div></td>
                                </tr>
                              </table>
                              <%end if%>
                            </td>
                            <td width="180" height="130"> 
                              <%if commended_pros>=3 then%>
                              <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
                                <tr> 
                                  <td> <a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(3)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                    <img src="proimages/<%=photo(3)%>" width="120" height="104" border="0"> 
                                    </a> </td>
                                </tr>
                                <tr> 
                                  <td height="25"> <div align="center"> <font size="2"> 
                                      <a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(3)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                      <%=modelno(3)%> </a> </font> </div></td>
                                </tr>
                              </table>
                              <%end if%>
                            </td>
                            <td width="181" height="130"> 
                              <%if commended_pros>=4 then%>
                              <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
                                <tr> 
                                  <td> <a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(4)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                    <img src="proimages/<%=photo(4)%>" width="120" height="104" border="0"> 
                                    </A> </td>
                                </tr>
                                <tr> 
                                  <td height="25"> <div align="center"> <font size="2"> 
                                      <a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(4)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                      <%=modelno(4)%> </A> </font> </div></td>
                                </tr>
                              </table>
                              <%end if%>
                            </td>
                          </tr>
                        </table> </td>
                    </tr>
                  </table> </td>
                <td height="563" valign="top" background="star/dd_boot_22.jpg">
				<img src="star/dd_boot_2.jpg" width="9" height="23"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="1000" border="0" cellspacing="0" cellpadding="0" height="10">
  <tr> 
    <td background="star/dd_bor_1.jpg" height="10"></td>
  </tr>
</table>
<table width="1000" border="0" cellspacing="0" cellpadding="0" height="91">
  <tr> 
    <td bgcolor="#EAEAEA" height="32">
	<div align="center">
	<!--#INCLUDE FILE="inc/bottom.htm" -->
	</div></td>
  </tr>
</table>
<%
conn.close
set conn=nothing
%>

<script language="JScript">
<!--
function form2submit()
{
  document.form2.submit();
}

function check1(form1)
{
if (form1.keywords.value=="")
 {
    alert("Input your keywords!")
    return false
  }
  else
  {
  return true
  }
}
//-->
</script>
</body>
</html>
