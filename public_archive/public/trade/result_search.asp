<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="conn.asp" -->
<%set rs=server.CreateObject("adodb.recordset")%>
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
	COLOR: #003399; TEXT-DECORATION: none
}
.A1:visited {
	COLOR: #003399; TEXT-DECORATION: none
}
.A1:active {
	COLOR: #990000; TEXT-DECORATION: none
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
    <td><!--#include file="inc/727.asp"-->
      <table width="769" height="36" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="18"><table width="768" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><TABLE height=72 cellSpacing=0 borderColorDark=#ffffff cellPadding=0 
            width="100%" borderColorLight=#666699 border=1>
                    <TBODY>
                      <TR> 
                        <FORM action="search.asp" method=post target=_blank>
                          <TD align=middle bgColor=#f5f5f5 height=66> <TABLE height=55 cellSpacing=0 cellPadding=0 width="100%" 
                  border=0>
                              <TBODY>
                                <TR> 
                                  <TD borderColor=#9ccfff align=middle width="81%" 
height=27> <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                                      <TBODY>
                                        <TR> 
                                          <TD width=171> <P align=right> <FONT color=#2975b5 
                              size=3><STRONG>Search:</STRONG></FONT><STRONG><FONT 
                              color=#ff0000>&nbsp;</FONT></STRONG></P></TD>
                                          <TD width=584> <TABLE cellSpacing=0 cellPadding=0 width="97%" 
                              border=0>
                                              <TBODY>
                                                <TR> 
                                                  <TD width="55%"> <DIV align=center> 
                                                      <INPUT id=key3 size=35 
										  onkeydown="if(event.keyCode==13) return check_search_empty(this.form)" 
										  name=key>
                                                    </DIV></TD>
                                                  <TD width="22%"> <P> 
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
                                                        <OPTION value="" selected>--All 
                                                        Products--</OPTION>
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
                                                    </P></TD>
                                                  <TD width="23%"> <DIV align=center> 
                                                      <INPUT class=list onclick="return check_search_empty(this.form);" type=submit value=Search name=Submit>
                                                    </DIV></TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE></TD>
                                          <TD width=27>&nbsp; </TD>
                                        </TR>
                                      </TBODY>
                                    </TABLE></TD>
                                </TR>
                                <TR> 
                                  <TD align=middle width="81%" height=28> 
                                    <%
					  sql="select id from products where recommended=true"
					  rs.open sql,conn,1,1
					  records=rs.recordcount
					  rs.close
					  if records<10 then
					    sql="select name,id,up_date from products where recommended=true order by up_date"
						 elseif records>=10 then
						sql="select top 10 name,id,up_date from products where recommended=true order by up_date"
					  end if
					  set rs=conn.execute(sql)
					  %>
                                    <TABLE height=0 cellSpacing=0 cellPadding=0 
width="100%">
                                      <TBODY>
                                        <TR> 
                                          <TD> <DIV align=center><font color="#000066"><strong>Hot 
                                              Products</strong></font> 
											  <%do while not rs.eof%>
											  &nbsp;
                                              <a href="pro_detail.asp?id=<%=rs("id")%>" class="A1"><font size="2"> 
                                              <%=rs("name")%>
											  </font>
											  </a>
											<%
					                           if not rs.eof then rs.movenext
					                           loop
					                           rs.close
					                         %>  
											  </DIV></TD>
                                        </TR>
                                      </TBODY>
                                    </TABLE>
									
									<%if records>10 then%>
                                    <table width="97%" border="0" cellspacing="0" cellpadding="0">
                                     
                                      <tr> 
                                        <td width="92%"><div align="right"><a href="pro_detail.asp"><font size="1"><b>more..</b></font></a></div></td>
                                      </tr>
                                    </table>
                                    <%
					  else
					  end if
					  %>
                                  </TD>
                                </TR>
                              </TBODY>
                            </TABLE></TD>
                        </FORM>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
              <tr>
                <td><table width="768" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="4">
                        <div align="center"><strong><font color="#FF9900">Sorry, 
                          there is no product that matches your keywords. Please 
                          try again or inform of us by click <font color="#000099"><a href="demand.asp">here</a></font> 
                          .</font></strong></div>
                      </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="18"> <TABLE height=159 cellSpacing=0 cellPadding=0 width="771" bgColor=#ffffff 
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
            </TABLE></td>
        </tr>
      </table></td>
  </tr>
</table>
<DIV align=center></DIV>
<DIV align=center></DIV>
<DIV align=center> 
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
		if(myform.pro_name.value == "")
		{
			alert("Please input your product name!");
			myform.pro_name.focus();
			return false;
		}
	}
</SCRIPT>
</BODY></HTML>
