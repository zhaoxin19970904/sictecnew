<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%response.Buffer=true%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Display picture in pages</title>
<link rel="stylesheet" href="upload.css" type="text/css">
<script language="JScript">
<!--
function confirmdel()
{
var fig=false
  fig=confirm("Delete all infor")
  return fig
}

function check2(form2)
{
var fig
 if (form2.proname.value=="")
 {
  fig=false
  window.alert("Please input the type you want to search (Type)")
  }
  return fig
}
function check1(form1)
{
var fig
if (form1.inputpage.value=="")
 {
  fig=false
  window.alert("Please input the page you want to go")
  }
  return fig
}
  //-->
</script>

</head>
<body>
<%
if session("admin_name")="" then 
  response.Redirect("../index.htm")
  response.End()
end if

dim rs,sql
dim iPage
set rs=server.CreateObject("adodb.recordset")
proname=request.form("proname")
if request.form("proname")<>"" then
sql="select * from products where type like '%"&request.form("proname")&"%' order by up_date desc"
 else
  sql="select * from products order by up_date desc"
end if
rs.pagesize=12
rs.open sql,conn,1,1
inputpage=request.form("inputpage")
if not isnumeric(inputpage) then 
response.write "NO results.Please input correct page��"
response.Write "<a href='#' onclick='javascript:history.go(-1)'>Back</a>"
response.End()
end if
if cint(inputpage)>cint(rs.pagecount) then 
response.write "NO results.Please input correct page��"
response.Write "<a href='#' onclick='javascript:history.go(-1)'>Back</a>"
response.End()
end if

if len(inputpage)<>0 then
   iPage=cint(inputpage)
   rs.absolutepage=inputpage
else   
if len(request("page"))=0 then
  iPage=1 
  else
   iPage=request("page")
 end if
  if not rs.eof then rs.absolutepage=iPage
end if 

%>
<p align="center" class="dat��eformat"><font color="#000066" size="4">�������ϴ��Ĳ�ƷͼƬ<br>
  <img src="../../images/line_01.jpg" width="338" height="20"> </font></p>
<p align="center" class="dateformat">��<font color="#FF0000">
  <%if len(inputpage)=0 and isnumeric(inputpage) then%>
  <%=iPage%> 
  <%else%>
  <%=inputpage%> 
  <%end if%></font>
  ҳ&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; �ܹ� <font color="#FF0000"><%=rs.recordcount%></font> �Ų�ƷͼƬ&nbsp; &nbsp;�ܹ� <font color="#FF0000"><%=rs.pagecount%></font> ҳ</p>


<table width="2780" height="58" border="0" align="center"  cellpadding="4" cellspacing="1" bgcolor="#293863">
  <tr bgcolor="#FFFFFF"> 
    <td width="63" class="chinese1"><div align="center">Title Type</div></td>
    <td width="63" height="29" class="chinese1"> <div align="center">Type</div></td>
    <td width="63" class="chinese1"> <div align="center">Name</div></td>
    <td width="163" class="chinese1"> <div align="center">Model</div></td>
    <td width="101" class="chinese1"> <div align="center">Brandname</div></td>
    <td width="53" class="chinese1"> <div align="center">Region</div></td>
    <td width="101" class="chinese1"> <div align="center">Approvals</div></td>
    <td width="82" class="chinese1"> <div align="center">Price</div></td>
    <td width="119" class="chinese1"> <div align="center">Main_export_markets</div></td>
    <td width="53" class="chinese1"> <div align="center">Packing</div></td>
    <td width="120" class="chinese1"> <div align="center">Key specification</div></td>
    <td width="82" class="chinese1"> <div align="center">Payment terms</div></td>
    <td width="44" class="chinese1"> <div align="center">Lead time</div></td>
    <td width="107" class="chinese1"> <div align="center">Cost</div></td>
    <td width="93" class="chinese1"><div align="center">Ssmple Available</div></td>
    <td width="93" class="chinese1"><div align="center">Related Products</div></td>
    <td width="93" class="chinese1"><div align="center">Key Words</div></td>
    <td width="93" class="chinese1"><div align="center">Manufacture</div></td>
    <td width="93" class="chinese1"><div align="center">Up Date</div></td>
    <td width="93" class="chinese1"> <div align="center">Up Mnanger</div></td>
    <td width="93" class="chinese1"> <div align="center">Porduct Photo</div></td>
    <td width="91" class="chinese1"> <div align="center">Note</div></td>
    <td width="389" class="chinese1"> <div align="center">��ز���</div></td>
  </tr>
  <%
  for i=1 to rs.pagesize
   if not rs.eof then
 %>
  <tr bgcolor="#fafbfd"> 
    <td class="chinese"><div align="center"><%=rs("title_type")%></div></td>
    <td height="26" class="chinese"> <div align="center"><%=rs("type")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("name")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("model")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("brandname")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("region")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("approvals")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("price")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("main_export_markets")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("packing")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("key_specification")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("payment_terms")%></div></td>
    <td  class="chinese"> <div align="center"><%=rs("lead_time")%></div></td>
    <td  class="chinese"> <div align="center"><%=rs("cost")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("sample")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("related_products")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("key_words")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("manufacture")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("up_date")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("up_admin")%></div></td>
    <td  class="chinese"> <div align="center"> <a href="prophotos/<%=rs("photo")%>" target="_blank">��ʾͼƬ</a> 
      </div></td>
    <td  class="chinese"><div align="center"> <a href="viewnote.asp?id=<%=rs("id")%>">�鿴����</a> 
      </div></td>
    <td  class="chinese"><div align="center"> <a href="endup.asp?id=<%=rs("id")%>" onClick="return confirmdel()">ɾ��</a> 
        | <a href="editproinfo.asp?id=<%=rs("id")%>">�޸Ĳ�Ʒ��Ϣ</a> 
        | <a href="editphoto.asp?id=<%=rs("id")%>">�޸Ĳ�ƷͼƬ</a> 
        | 
        <%if rs("recommended")=false then%>
        <a href="deal.asp?id=<%=rs("id")%>&action=tj">��Ϊ�Ƽ�</a> 
        <%else%>
        <a href="deal.asp?id=<%=rs("id")%>&action=cxtj"><font color="#ff0000">�����Ƽ�</font></a> 
        <%end if%>
        | 
        <%if rs("especially")=false then%>
        <a href="deal.asp?id=<%=rs("id")%>&action=ts">��Ϊ�����Ʒ</a> 
        <%else%>
        <a href="deal.asp?id=<%=rs("id")%>&action=cxts"><font color="#ff0000">���������Ʒ</font></a> 
        <%end if%>
        | 
        <%if rs("ifnew")=false then%>
        <a href="deal.asp?id=<%=rs("id")%>&action=isnew">��Ϊ���²�Ʒ</a> 
        <%else%>
        <a href="deal.asp?id=<%=rs("id")%>&action=cxnew"><font color="#ff0000">�������²�Ʒ</font></a> 
        <%end if%>
      </div></td>
  </tr>
  <%
 end if
 if not rs.eof then rs.movenext
 next
 %>
</table>
<p align="center"> 
  <%
if cint(iPage)=1 or cint(inputpage)=1 then
%>
  <font class="dateformat">��һҳ | ��һҳ | </font> 
  <%
else
%>
  <a href="viewpro.asp?page=1">��һҳ</a> | <a href="viewpro.asp?page=<%=iPage-1%>">��һҳ</a> 
  | 
  <%end if%>
  <%
 if cint(iPage)=cint(rs.pagecount) or cint(inputpage)=cint(rs.pagecount) then
%>
  <font class="dateformat">��һҳ | ���һҳ </font>
  <%else%>
  <a href="viewpro.asp?page=<%=iPage+1%>">��һҳ</a> | <a href="viewpro.asp?page=<%=rs.pagecount%>">���һҳ</a> 
  <%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  <br>
<form name="form1" method="post" action="" onsubmit="return check1(this)">
  <div align="center" class="chinese">ת���� 
    <input name="inputpage" type="text" size="3" class="input" style="ime-mode:disabled">
    ҳ 
    <input type="submit" name="Submit" value="GO">
  </div>
</form>

<form name="form2" method="post" action="viewpro.asp?ifsearch=1" onsubmit="return check2(this)">
  <table width="508" height="16" border="0" align="center" cellpadding="0" cellspacing="0" class="chinese">
    <tr>
      <td width="268" height="16">����Ʒ��(Type)������ 
        <input name="proname" type="text" id="proname" size="15" class="input" value=<%=proname%>></td>
      <td width="106"> <div align="left">
          <input type="submit" name="Submit2" value="��ʼ����" class="button">
          &nbsp;&nbsp;&nbsp;&nbsp;|</div></td>
      <td width="120"><input type="button" name="Submit3" value="��ʾȫ��" class="button" onClick="javascript:location.href='viewpro.asp'"></td>
      <td width="14">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
