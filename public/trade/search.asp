<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="conn.asp" -->
<!--#include file="script/search.asp" -->
<HTML><HEAD><TITLE>Sictec Product & Manufacturer Catalogue:  <%=type_type%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META content="China Online Commodities Fair" name=title>
<META 
content="china products; china manufacturers; chinese manufacturers; china product; china suppliers; doing business with china; china supplier; chinese product; trading with china; products of china; business with china; products made in china; made in china" 
name=keywords>
<META 
content="ningbo,export,trade,the leading global eletronic commerce site, offers virtual trade point web services such as Trade Leads, Product Catalogs, Company Directories." 
name=description>
<STYLE type=text/css>
A:link {
	COLOR: #003399; TEXT-DECORATION: underline
}
A:visited {
	COLOR: #003399; TEXT-DECORATION: underline
}
A:active {
	COLOR: #990000; TEXT-DECORATION: underline
}
A:hover {
	COLOR: #990000; TEXT-DECORATION: underline
}



.A1:link {
	COLOR: #2c65aa; TEXT-DECORATION: none
}
.A1:visited {
	COLOR: #2c65aa; TEXT-DECORATION: none
}
.A1:active {
	COLOR: #990000; TEXT-DECORATION: none
}
.A1:hover {
	COLOR: #990000; TEXT-DECORATION: none
}

.A2:link {
	COLOR: #626365; TEXT-DECORATION: none
}
.A2:visited {
	COLOR: #626365; TEXT-DECORATION: none
}
.A2:active {
	COLOR: #e43b00; TEXT-DECORATION: none
}
.A2:hover {
	COLOR: #e43b00; TEXT-DECORATION: underline
}


.coverer A:link {
	COLOR: #ffffff; TEXT-DECORATION: underline
}
.coverer A:visited {
	COLOR: #ffffff; TEXT-DECORATION: underline
}
.coverer A:active {
	COLOR: #ffff00; TEXT-DECORATION: none
}
.coverer A:hover {
	COLOR: #ffff00; TEXT-DECORATION: none
}
BODY {
	FONT-SIZE: 11px; COLOR: #000000; FONT-FAMILY: "Verdana","Arial", "Helvetica"
}
TD {
	FONT-SIZE: 12px; FONT-FAMILY: "Verdana","Arial", "Helvetica"; TEXT-DECORATION: none
}
.fonten {
	FONT-SIZE: 11px; FONT-FAMILY: "Arial", "Helvetica"; TEXT-DECORATION: none
}
.list {
	FONT-SIZE: 12px; FONT-FAMILY: "Arial", "Helvetica"; TEXT-DECORATION: none
}
.line {
	LINE-HEIGHT: 20px
}
.big {
	FONT-SIZE: 17px; FONT-FAMILY: "Arial", "Helvetica"; TEXT-DECORATION: none
}
INPUT {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 15px; FONT-FAMILY: Tahoma,Verdana
}
SELECT {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 15px; FONT-FAMILY: Tahoma,Verdana
}
TEXTAREA {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 15px; FONT-FAMILY: Tahoma,Verdana
}
OPTION {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 15px; FONT-FAMILY: Tahoma,Verdana
}
.stdedit {
	BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #666666 1px solid; FONT-SIZE: 12px; BORDER-LEFT: #666666 1px solid; BORDER-BOTTOM: #cccccc 1px solid; FONT-FAMILY: "Verdana", "Arial", "Helvetica", "sans-serif"; BACKGROUND-COLOR: #ffffff
}
.stdedit1 {
	BORDER-RIGHT: #666666 1px solid; BORDER-TOP: #cccccc 1px solid; FONT-SIZE: 11px; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #666666 1px solid; FONT-FAMILY: "Arial", "Helvetica", "sans-serif"; BACKGROUND-COLOR: #eeeeee
}
</STYLE>

<META content="MSHTML 6.00.2900.2668" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 bgcolor="#FFFFFF">
<table width="150" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td><!--#include file="inc/727.asp"--></td>
  </tr>
</table>
<DIV align=center></DIV>
<DIV align=center></DIV>
<DIV align=center> 
  <TABLE height=467 cellSpacing=0 cellPadding=0 width=770 border=0>
    <TBODY> 
    <TR> 
      <TD vAlign=top width=184><BR>
        <TABLE cellSpacing=1 cellPadding=0 width="100%" bgColor=#6192c4 
        border=0>
          <TBODY> 
          <TR> 
                <TD width="100%" bgColor=#6192c4 borderColorDark=#6192c4 
            height=22>&nbsp;<B><font color="#FFFFFF">Product&nbsp; search</font></B></TD>
          </TR>
          <TR> 
            <TD width="100%" bgColor=#ffffff>
			 <% set rs=server.createobject("adodb.recordset")%>
              <FORM name="form1" action="search.asp" method="post">
                    <TABLE height=93 cellSpacing=0 width="100%">
                      <TBODY>
                        <TR> 
                          <TD width="18%" bgColor=#f7fafd height=37> 
                            <DIV align=right><img src="images/search.jpg" width="13" height="13"><BR>
                            </DIV></TD>
                          <TD width="82%" height=37 valign="middle" bgColor=#f7fafd> 
                            <div align="center"><FONT color=#000000> 
                              <INPUT id=key size=20 
										  onkeydown="if(event.keyCode==13) return check_search_empty(this.form)" 
										  name=key>
                              </FONT></div></TD>
                        </TR>
                        <TR> 
                          <TD height=25 colspan="2" bgColor=#f7fafd> <DIV align=right></DIV>
                            <div align="center"><FONT color=#000000> 
							<%
										sql="select title_id from title"
										rs.open sql,conn,1,1
										title_reds=rs.recordcount
										rs.close
										dim title_id(),i,title_type()
										i=1
										if title_reds>=12 then
										 sql="select top 12 * from title"
										    elseif title_reds<12 then
										 sql="select * from title"
										end if
										rs.open sql,conn,1,1
										title_reds=rs.recordcount
										%>
                              <SELECT name=title_type class=list id="title_type">
                                <OPTION value="" selected>--All Products--</OPTION>
                                <%
								do while not rs.eof
								redim preserve title_id(i),title_type(i)
								title_type(i)=rs("title_type")
								title_id(i)=rs("title_id")
								%>
                                <OPTION value="<%=title_type(i)%>"><%=title_type(i)%></OPTION>
                                <%
								if not rs.eof then rs.movenext
								i=i+1
								loop
								rs.close
								%>
                              </SELECT>
                              </FONT></div></TD>
                        </TR>
                        <TR> 
                          <TD vAlign=center bgColor=#f7fafd colSpan=2 height=1> 
                           </TD>
                        </TR>
                        <TR> 
                          <TD vAlign=center bgColor=#f7fafd colSpan=2 height=24> 
                            <DIV align=center> 
                              <INPUT class=list onclick="return check_search_empty(this.form);" type=image
		    src="star/But_GO.gif" width="59" height="19">
                            </DIV></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
              </FORM>
            </TD>
          </TR>
          </TBODY>
        </TABLE>
        
          <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY> 
          <TR> 
            <TD width="100%" height=6></TD>
          </TR>
          </TBODY>
        </TABLE>
          <TABLE height=292 cellSpacing=1 cellPadding=0 width="100%" bgColor=#6192c4 
      border=0>
            <TBODY> 
          <TR> 
                <TD vAlign=top width="100%" bgColor=#ffffff borderColorDark=#6192c4 
          height=277> 
                  <TABLE height=40 cellSpacing=0 cellPadding=0 width="100%" 
              border=0>
                <TBODY> 
                <TR> 
                  <TD width="100%" bgColor=#e4edf9 height=24><B>Hot sale</B></TD>
                </TR>
                <TR> 
                  <TD width="200" height=16><div align="center"><FONT 
                  color=#00309c>&nbsp;</FONT><B><FONT color=#ff0000>[ Sale ]</FONT></B></div></TD>
                </TR>
                </TBODY>
              </TABLE>
              <DIV align=center> 
                <CENTER>
				      <%
					  sql="select id from products where recommended=true"
					  rs.open sql,conn,1,1
					  records=rs.recordcount
					  rs.close
					  if records<20 then
					    sql="select name,id,up_date from products where recommended=true order by up_date"
						 elseif records>=20 then
						sql="select top 20 name,id,up_date from products where recommended=true order by up_date"
					  end if
					  set rs=conn.execute(sql)
					  %>
                      <table width="97%" border="0" cellspacing="0" cellpadding="0">
                        <%
						do while not rs.eof
						up_date=split(rs("up_date")," ")
						%>
                        <tr>
						  <td width="8%" height="18" align="center" valign="top"><strong>.</strong></td>
                          <td width="92%"><a href="pro_detail.asp?id=<%=rs("id")%>"><font size="2"><%=rs("name")%></font><font size="1"><b>(<%=up_date(0)%>)</b></font></a></td>
                        </tr>
					<%
					if not rs.eof then rs.movenext
					loop
					rs.close
					%>
                      </table>
					  <%if records>20 then%>
					  <table width="97%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
						  <td width="8%" height="14" align="center" valign="middle">&nbsp;</td>
						 </tr>
						 <tr>
                          <td width="92%"><div align="right"><a href="more_hot_pro.asp"><font size="1"><b>more..</b></font></a></div></td>
                        </tr>
                      </table>
					  <%
					  else
					  end if
					  %>
                    </CENTER>
              </DIV>
            </TD>
          </TR>
          </TBODY>
        </TABLE>
      </TD>
      <TD width=7> 
        <TD width=4 height=467></TD>
        <TD vAlign=top align=middle width=575><br>
		
       
                  <SCRIPT language=JavaScript>
<!--
function Trim(str){
	if(str.charAt(0) == ' '){
		 str = str.slice(1);
		str = Trim(str); 
	}
	return str;
}

function check_search_empty(myform){

							var search=myform.key.value;
                           
								  if(Trim(search)=='')
								  {

								     myform.key.focus();

									alert('Please input search keywords!');

									 return false;

							      }  
}
//-->
</SCRIPT>
		   <%
		   key1=request.Form("key")
		   sql=search()
		   'key=request.Form("key")
		   'title_type=request.Form("title_type")
		'sql="select id,type,photo,name,model,type_id from products where type_id="&type_id&" order by up_date desc"
		rs.open sql,conn,1,1
			'if rs.eof then
	           'response.Redirect("result_search.asp")
               'response.end
            'end if

if len(request("page"))=0 then
  iPage=1 
  else
   iPage=request("page")
 end if

 rs.pagesize=10
 rs.absolutepage=iPage
 rs.cachesize=rs.pagesize
  i=0
dim id(),pro_photo(),p_name(),model(),typeid()
type_type=rs("type")
do while not rs.eof and (i<rs.pagesize)
          i=i+1
     redim preserve id(i),pro_photo(i),p_name(i),model(i),typeid(i)
          id(i)=rs("id")
		  p_name(i)=rs("name")
		  typeid(i)=rs("type_id")
          pro_photo(i)=rs("photo")
		  model(i)=rs("model")
		  rs.movenext
loop 
					 %>
        <TABLE height=291 cellSpacing=0 cellPadding=0 width="100%" bgColor=#ffffff border=0>
            <TBODY> 
          <TR> 
                <TD width=573 height=18> <font color="#000099"><strong>Search 
                  </strong>&nbsp;"<font color="#ff9900"><strong><%=key1%></strong></font>"<strong> &nbsp;result :</strong></font> </TD>
          </TR>
          <TR> 
                <TD vAlign=top width=573 height=273><table width="573" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td height="7"></td>
                    </tr>
                    <tr> 
                      <td height="20"> 
                        <table width="571" height="19" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="285" height="19"><font color="#0000CC">Total: 
                              <%=rs.recordcount%> Products</font></td>
                            <td width="286"><div align="right"><font color="#0000CC"> 
                                <%
if cint(iPage)=1 or cint(inputpage)=1 then
else
%>
                                <a href="showroomview.asp?page=<%=iPage-1%>"><b>Pre</b></a>&nbsp; 
                                <%end if%>
                                <%if len(inputpage)=0 and isnumeric(inputpage) then%>
                                <%=iPage%> 
                                <%else%>
                                <%=inputpage%> 
                                <%end if%>
                                / <%=rs.pagecount%>
                                <%
if cint(iPage)=cint(rs.pagecount) or cint(inputpage)=cint(rs.pagecount) then
else
%>
                                <a href="showroomview.asp?page=<%=iPage+1%>">&nbsp;<b>Next</b></a> 
<%
end if
%>
                                &nbsp;&nbsp;&nbsp;</font></div></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td height="220" valign="top"> 
<%
					 'sql="select id,photo,name,model from products where type_id="&type_id&" order by up_date desc"
					 'rs.open sql,conn,1,1
	 
i=1 
Do while i<=ubound(id)
%>
                        <table width="100" border="0" cellspacing="0" cellpadding="0" align="center">
                          <tr> 
                            <td width="500" height="13"></td>
                          </tr>
                        </table> 
						
                        <table width="572" height="20" border="0" cellpadding="0" cellspacing="0">
                          <tr bgcolor="#aca899"> 
                            <td colspan="3" height="1"></td>
                          </tr>
						  
						  <tr> 
                            <td colspan="3" height="8"></td>
                          </tr>
						  
                          <tr> 
                            <td width="279" height="18"><table width="277" height="19" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td width="110" height="19"> <div align="center"><a href="pro_detail.asp?id=<%=id(i)%>" target="_blank"> 
                                      <img src="admin/script/prophotos/<%=pro_photo(i)%>" border="0" height="92" width="92"> 
                                      </a> </div></td>
                                  <td width="167"><table width="167" height="92" border="0" cellpadding="0" cellspacing="0">
                                      <tr> 
                                        <td height="18" colspan="3"> <img src="images/pic.jpg" width="11" height="11"> 
                                          <a href="pro_detail.asp?id=<%=id(i)%>" target="_blank"><%=model(i)%></a></td>
                                      </tr>
                                      <tr> 
                                        <td height="56" colspan="2"> <font size="2" color="#666666"><%=type_type%> (<%=p_name(i)%>)</font> </td>
                                      </tr>
                                      <tr> 
                                        <td width="95" height="18">
                                      <div align="center"><font size="2"><b> <a href="script/basket.asp?pro_id=<%=id(i)%>&type_id=<%=typeid(i)%>" class="A2" target="_blank"><img src="images/basket.jpg" width="84" height="14" border="0"></a></b></font></div>
                                    </td>
                                        <td width="72">
                                      <div align="center"> <a href="inquiry.asp?pro_id=<%=id(i)%>" target="_blank"> 
                                        <img src="images/inqurie.jpg" width="70" height="14" border="0"> 
                                        </a> </div>
                                    </td>
                                      </tr>
                                    </table></td>
                                </tr>
                              </table></td>
							<TD width="14">&nbsp;</TD>
                            <td width="279">
			<%
			i=i+1
			if i<=ubound(id) then 
			%>
							<table width="277" height="19" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="110" height="19"><div align="center"><a href="pro_detail.asp?id=<%=id(i)%>" target="_blank"><img src="admin/script/prophotos/<%=pro_photo(i)%>" border="0" height="92" width="92"></a></div></td>
                                  <td width="167">
                                <table width="167" height="91" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height="18" colspan="2"><img src="images/pic.jpg" width="11" height="11"> 
                                      <a href="pro_detail.asp?id=<%=id(i)%>" target="_blank"><%=model(i)%></a> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td height="55" colspan="2"><font size="2" color="#666666"><%=type_type%> 
                                      (<%=p_name(i)%>)</font></td>
                                  </tr>
                                  <tr> 
                                    <td width="95" height="18"> 
                                      <div align="center"><font size="2"><b> <a href="script/basket.asp?pro_id=<%=id(i)%>&type_id=<%=typeid(i)%>" class="A2" target="_blank"><img src="images/basket.jpg" width="84" height="14" border="0"></a></b></font></div>
                                    </td>
                                    <td width="72"> 
                                      <div align="center"> <a href="inquiry.asp?pro_id=<%=id(i)%>" target="_blank"> 
                                        <img src="images/inqurie.jpg" width="70" height="14" border="0"> 
                                        </a> </div>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                                </tr>
                              </table>
			<% 
              i=i+1
              end if
            %>	
							  </td>
                          </tr>
						  <tr> 
                            <td colspan="3" height="3"></td>
                          </tr>
                          <tr bgcolor="#aca899"> 
                            <td colspan="3" height="1"></td>
                          </tr>
                        </table>
						
<%
loop
%>
                      </td>
                    </tr>
                    <tr> 
                      <td><table width="571" height="18" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            
                        <td width="285" height="18"><font color="#0000CC">Total: 
                          <%=rs.recordcount%> Products</font></td>
                            <td width="286"><div align="right">
                                <div align="right"><font color="#0000CC"> 
                                  <%
if cint(iPage)=1 or cint(inputpage)=1 then
else
%>
                                  <a href="index.asp?page=<%=iPage-1%>"><b>Pre</b></a>&nbsp; 
                                  <%end if%>
                                  <%if len(inputpage)=0 and isnumeric(inputpage) then%>
                                  <%=iPage%> 
                                  <%else%>
                                  <%=inputpage%> 
                                  <%end if%>
                                  / <%=rs.pagecount%> 
                                  <%
 if cint(iPage)=cint(rs.pagecount) or cint(inputpage)=cint(rs.pagecount) then
else%>
                                  <a href="index.asp?page=<%=iPage+1%>">&nbsp;<b>Next</b></a> 
                                  <%
end if
%>
                                  &nbsp;&nbsp;&nbsp;</font></div>
                              </div></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table>
				  <%rs.close%>
                  <font color="#000066">&nbsp;</font> </TD>
          </TR>
          </TBODY>
        </TABLE>
        
        </TD>
    </TR>
    </TBODY>
  </TABLE>
  
  <TABLE height=159 cellSpacing=0 cellPadding=0 width="771" bgColor=#ffffff 
      border=0>
    <TBODY>
      <TR> 
        <TD width=771 height=24><IMG height=23 
            src="star/featured_pro.gif" 
            border=0></TD>
      </TR>
      <TR> 
        <%
		  dim photo(),pro_name(),pro_id(),j
		  j=1
		  sql="select id from products where especially=true"
		  rs.open sql,conn,1,1
		  sep_pro_reds=rs.recordcount
		  if sep_pro_reds>=6 then
		  sql="select top 6 id,name,photo from products where especially=true"
		  else
		  sql="select id,name,photo from products where especially=true"
		  end if
		  rs.close
		  set rs=conn.execute(sql)
		  do while not rs.eof
		    redim preserve photo(j),pro_name(j),pro_id(j)
			photo(j)=rs("photo")
			pro_name(j)=rs("name")
			pro_id(j)=rs("id")
			if not rs.eof then rs.movenext
			j=j+1
		  loop
		  rs.close
		  arrary_count=ubound(photo)
		  %>
        <TD width=771 height="135"> <DIV align=center> 
            <TABLE height=133 width="771">
              <TBODY>
                <TR> 
                  <TD vAlign=top align=middle width="180" height=129> 
                    <%if arrary_count>=1 then%>
                    <div align="center"> <a href="pro_detail.asp?id=<%=pro_id(1)%>"><IMG height=84 alt="<%=pro_name(1)%>" 
                  src="admin/script/prophotos/<%=photo(1)%>" 
                  width=99 align=center vspace=4 border=0></a><BR>
                      <a href="pro_detail.asp?id=<%=pro_id(1)%>"><%=pro_name(1)%></a> 
                    </div>
                    <%end if%>
                  </TD>
                  <TD vAlign=top align=middle width="180" height=129> 
                    <%if arrary_count>=2 then%>
                    <div align="center"> <a href="pro_detail.asp?id=<%=pro_id(2)%>"><IMG height=84 alt="<%=pro_name(2)%>" 
                  src="admin/script/prophotos/<%=photo(2)%>" 
                  width=99 align=center vspace=4 border=0></a><BR>
                      <a href="pro_detail.asp?id=<%=pro_id(2)%>"><%=pro_name(2)%></a> 
                    </div>
                    <%end if%>
                  </TD>
                  <TD vAlign=top align=middle width="180" height=129> 
                    <%if arrary_count>=3 then%>
                    <div align="center"> <a href="pro_detail.asp?id=<%=pro_id(3)%>"><IMG height=84 
                  alt="<%=pro_name(3)%>" 
                  src="admin/script/prophotos/<%=photo(3)%>" 
                  width=99 align=center vspace=4 border=0></a><BR>
                      <a href="pro_detail.asp?id=<%=pro_id(3)%>"><%=pro_name(3)%></a> 
                    </div>
                    <%end if%>
                  </TD>
                  <TD vAlign=top align=middle width="180" height=129> 
                    <%if arrary_count>=4 then%>
                    <div align="center"> <a href="pro_detail.asp?id=<%=pro_id(4)%>"><IMG height=84 alt="<%=pro_name(4)%>" 
                  src="admin/script/prophotos/<%=photo(4)%>" 
                  width=99 align=center vspace=4 border=0></a><BR>
                      <a href="pro_detail.asp?id=<%=pro_id(4)%>"><%=pro_name(4)%></a> 
                    </div>
                    <%end if%>
                  </TD>
                  <TD vAlign=top align=middle width="180" height=129> 
                    <%if arrary_count>=5 then%>
                    <div align="center"> <a href="pro_detail.asp?id=<%=pro_id(5)%>"><IMG height=84 alt="<%=pro_name(5)%>" 
                  src="admin/script/prophotos/<%=photo(5)%>" 
                  width=99 align=center vspace=4 border=0></a><BR>
                      <a href="pro_detail.asp?id=<%=pro_id(5)%>"><%=pro_name(5)%></a> 
                    </div>
                    <%end if%>
                  </TD>
                  <TD vAlign=top align=middle width="180" height=129> 
                    <%if arrary_count>=6 then%>
                    <div align="center"> <a href="pro_detail.asp?id=<%=pro_id(6)%>"><IMG height=84 alt="<%=pro_name(6)%>" 
                  src="admin/script/prophotos/<%=photo(6)%>" 
                  width=99 align=center vspace=4 border=0></a><BR>
                      <a href="pro_detail.asp?id=<%=pro_id(6)%>"><%=pro_name(6)%></a> 
                    </div>
                    <%end if%>
                  </TD>
                </TR>
              </TBODY>
            </TABLE>
          </DIV></TD>
      </TR>
    </TBODY>
  </TABLE>
  <table width="150" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td> 
        <!--#include file="inc/919.asp"-->
      </td>
    </tr>
  </table>
  <%
  conn.close
  set conn=nothing
  %>
</DIV>
<DIV align=center> </DIV>
<SCRIPT language=JavaScript>
	function checkempty(myform)
	{
		if(myform.subject.value == "")
		{
			alert("Please input inquiry subject!");
			myform.subject.focus();
			return false;
		}
	}
</SCRIPT>
</BODY></HTML>
