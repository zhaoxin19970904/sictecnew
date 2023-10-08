<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="conn.asp" -->
<!-- saved from url=(0028)http://en.ningbo-export.com/ -->
<HTML><HEAD><TITLE>Sictec International Trade Co.,Ltd.</TITLE>
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
	COLOR: #666666; TEXT-DECORATION: underline
}
.A1:visited {
	COLOR: #666666; TEXT-DECORATION: underline
}
.A1:active {
	COLOR: #666666; TEXT-DECORATION: underline
}
.A1:hover {
	COLOR: #990000; TEXT-DECORATION: underline
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
  <TABLE height=454 cellSpacing=0 cellPadding=0 width=770 border=0>
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
          <TABLE height=278 cellSpacing=1 cellPadding=0 width="100%" bgColor=#6192c4 
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
						  <td width="8%" align="center" valign="top"><strong>.</strong></td>
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
        <TD width=4 height=454></TD>
        <TD vAlign=top align=middle width=575><BR>
		<%
		id=request.QueryString("id")
		sql="select * from title where title_id="&id&""
		set rs=conn.execute(sql)
		%>
        <TABLE height=20 cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY> 
          <TR> 
                <TD width="100%" height=20> <font color="#666666" size="2">You 
                  are here:</font> <a href="index.asp"><font size="2"><b>Product 
                  catalogs</b></font></a> <font color="#000066">></font> export 
                  <font color="#000066">></font> <strong>
				  <font color="#000099"><%=rs("title_type")%></font>
				  </strong><%rs.close%> 
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
                </TD>
          </TR>
          </TBODY>
        </TABLE>
        <TABLE height=180 cellSpacing=0 cellPadding=0 width="99%" bgColor=#ffffff 
      border=0>
          <TBODY> 
          <TR> 
            <TD width=562 height=24><IMG height=23 
            src="star/pro_cata.gif" width=564 
            border=0></TD>
          </TR>
          <TR> 
            <TD vAlign=top width=562 height=273> 
			<%
			 dim main_id(),m,main_type(),main_id_count
			 m=1
			 'id=request.QueryString("id")   'title_id
			 sql="select * from main where title_id="&id&""
			 rs.open sql,conn,1,1
			 
			 if rs.eof or rs.bof then
			 response.Redirect("index.asp")
			 response.End()
			 end if
			 
			 do while not rs.eof
			 redim preserve main_id(m),main_type(m)
			   main_id(m)=rs("main_id")
			   main_type(m)=rs("main_type")
			   if not rs.eof then rs.movenext
			    m=m+1
			 loop
			 rs.close
			 lows=0
			%>
                  <table width="567" height="22" border="0" cellpadding="0" cellspacing="0">
                    <tr>
					  <td width=12></td>
                      <td width="283" height="19" valign="top"> 
                        <table width="280" border="0" cellspacing="0" cellpadding="0">
						
		<%for main_id_count=1 to ubound(main_id)%>
		
                          <tr>
                            <td width="280" height="20"> 
							<font color="#666666" size="2"><b><font color="#333399"> 
                              <%=main_type(main_id_count)%>
							</font></b></font>
							</td>
                          </tr>
						  <%
						  set rs1=server.CreateObject("adodb.recordset")
						  sql="select * from type where main_id="&main_id(main_id_count)&""
						  set rs=conn.execute(sql)
						  'rs.open sql,conn,1,1
						  'main_id_records=rs.recordcount
						  do while not rs.eof
						    type_id=rs("type_id")
						    sql="select id from products where type_id="&type_id&""
							rs1.open sql,conn,1,1
							main_id_records=rs1.recordcount
						  %>
                          <tr>
                            <td width="280" height="20">
							&nbsp;
							<A href="showroomview.asp?type_id=<%=type_id%>&title_id=<%=id%>&main_id=<%=main_id(main_id_count)%>" class="A1">
							<font face="Geneva, Arial, Helvetica, san-serif">
							<%=rs("type")%>(&nbsp;<%=main_id_records%>&nbsp;products)
							</font>
							</A>
							</td>
                          </tr>
						  <%
						  if not rs.eof then rs.movenext
						  rs1.close
						  loop
						  rs.close
						  %>
						  
		<%next%>
						  
                        </table>
                      </td>
					  <td width=14></td>
                      <td width="258" valign="top">
                    <table width="256" border="0" cellspacing="0" cellpadding="0">
					
					<%%>
                          <tr> 
                            <td width="257">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td width="257">&nbsp;</td>
                          </tr>
					<%%>	  
						  
                        </table>
                      </td>
                    </tr>
                  </table> </TD>
          </TR>
          </TBODY>
        </TABLE>
        <br>
        <BR>
        <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY> 
          <TR> 
            <TD width="100%"></TD>
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
		if(myform.username.value == "")
		{
			alert("Please input your UserID!");
			myform.username.focus();
			return false;
		}
		
		if(myform.userpass.value == "")
		{
			alert("Please input your Password!");
			myform.userpass.focus();
			return false;
		}
		return true;
	}
</SCRIPT>
</BODY></HTML>
