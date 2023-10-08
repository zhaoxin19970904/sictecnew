<!--#include file="admin_conn.asp" -->
<%
'=========================================================
' File: index.asp
' Version:1.0.1
' Date: 2003-8-23
' Script Written by cheng
'=========================================================
' Copyright (C) 2002-2003 mtvok.com. All rights reserved.
' Web: http://www.mtvok.com,http://www.dvking.cn,http://www.dvonlin.cn
' Email: zcl22@21cn.com
'=========================================================

dim czlb,ltmess

if request.form("form1")<>"" and request.form("czlb")<>"" then
dim rs,sqltext
czlb=request.Form("czlb")
select case(czlb)
case("addlb")
 call addlbname()
case("editlb")
 call addlbname()
case ("addzlt")
 call addlbzlt()
case ("editzlb")
 call addlbzlt()
end select
end if

sub addlbname()
set rs=server.createobject("adodb.recordset")
if czlb="addlb" then
rs.Open "ltname", conn, 1, 3
rs.addnew
else
sqltext="select * from ltname where id="&request.Form("id")
rs.open sqltext,conn,1,3
end if
rs("ltname")=request.form("ltname")
rs("time")=now()
rs("b_order")=request.Form("b_order")
rs.update
ltmess=ltmess&"<li>操作成功!"
end sub

sub addlbzlt()
dim seting
set rs=server.createobject("adodb.recordset")
if request.Form("ltlb")="" or request.Form("ltsm")="" then 
ltmess=ltmess&"<li>本次操作不成功!<li>请填写好表单.再提交!<li>论坛名称,论坛说明是必项."
exit sub
end if
if czlb="addzlt" then
rs.Open "ltlb", conn, 1, 3
rs.addnew
seting="0|0|1|1|"
rs("today")=date()
rs("todayft")=0
else
seting=(request.Form("set1")&"|"&request.Form("set2")&"|"&request.Form("set3")&"|"&request.Form("set4"))
sqltext="select * from ltlb where id="&request.Form("id")
rs.Open sqltext, conn, 1, 3
end if

rs("ltlb")=request.form("ltlb")
rs("ltsm")=request.form("ltsm")
rs("face")=request.form("face")
rs("atlb")=request.Form("ltlbid")
rs("b_order")=request.Form("b_order")
rs("time")=now()
rs("ltconfig")=seting
rs.update
ltmess=ltmess&"<li>操作成功!"
end sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>论坛管理</title>
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
</head>

<body>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
  <tr bgcolor="#F7F7F7"> 
    <td colspan="3"><%=ltmess%></td>
  </tr>
  <%
call listzlt()
sub listzlt()
dim rs,sqltext,i
set rs=server.createobject("adodb.recordset")
sqltext="select ltname,ifoff,id from ltname order by b_order desc,id desc"
rs.open sqltext,conn,1,1
if not rs.eof then
for i=1 to rs.recordcount
%>
  <tr> 
    <td width="24%" height="39" bgcolor="#FFFFFF"><%=rs("ltname")%></td>
    <td width="38%" bgcolor="#FFFFFF"> 
      <%call listboard(rs("id"))%></td>
    <td width="38%" bgcolor="#FFFFFF"><a href="admin_cz.asp?czlb=dellb&id=<%=rs("id")%>" onclick="{if(confirm('确实要将该主论坛删除吗，确定删除吗?')){return true;}return false;}">删除</a> 
      <a href="admin_index.asp?czlb=editlb&id=<%=rs("id")%>">修改</a> 
    </td>
  </tr>
  <%
rs.movenext
next
end if 
end sub
%>

  <tr> 
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td colspan="2" bgcolor="#FFFFFF"><a href="admin_index.asp?czlb=addlb">添加类别</a> 
      <a href="admin_index.asp">复位 </a><a href="admin_index.asp?czlb=addzlt">添加子论坛</a></td>
  </tr>
</table>
<form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>">
  <div id="theplaylist" align="center"></div>
<%
czlb=request.QueryString("czlb")
select case(czlb)
case("editlb")
set rs=conn.execute("select ltname,ifoff,id,b_order from ltname where id="&request.QueryString("id"))
%>
  <table align="center" cellspacing="1" cellpadding="4" width="450" bgcolor="#0053A6">
    <tr align="center" bgcolor="#006699"> 
      <td colspan="2"><strong><font color="#FFFFFF">修改主分类别</font></strong></td>
    </tr>
    <tr> 
      <td width="128" bgcolor="#FFFFFF">类别名称 
        <input name="czlb" type="hidden" id="czlb" value="<%=czlb%>">
        <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td>
      <td width="301" bgcolor="#FFFFFF"> <input name="ltname" type="text" id="ltname" value="<%=rs("ltname")%>" size="30" maxlength="30"> 
      </td>
    </tr>
    <tr>
      <td width="128" bgcolor="#FFFFFF">论坛序号(数字)</td>
      <td bgcolor="#FFFFFF"><input name="b_order" type="text" id="b_order" value="<%=rs("b_order")%>">
        -- 用于排序,从大到小</td>
    </tr>
    <tr>
      <td width="128" bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF"> <input name="form1" type="submit" id="form1" value="修 改"></td>
    </tr>
  </table>
 <%
 case("addlb")
 %>
  <table align="center" cellspacing="1" cellpadding="4" width="450" bgcolor="#003366">
    <tr align="center" bgcolor="#006699"> 
      <td colspan="2"><strong><font color="#FFFFFF">添加主分类</font></strong></td>
    </tr>
    <tr> 
      <td width="128" bgcolor="#FFFFFF">类别名称 
        <input name="czlb" type="hidden" id="czlb2" value="<%=czlb%>"></td>
      <td width="301" bgcolor="#FFFFFF"> <input name="ltname" type="text" id="ltname" size="30" maxlength="30"> 
      </td>
    </tr>
    <tr>
      <td width="128" bgcolor="#FFFFFF">论坛序号(数字)</td>
      <td bgcolor="#FFFFFF"><input name="b_order" type="text" id="b_order">
        -- 用于排序,从大到小</td>
    </tr>
    <tr>
      <td width="128" bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF"> <input name="form1" type="submit" id="form1" value="添 加"></td>
    </tr>
  </table>
<%
case "addzlt"
%></p>
  <table width="59%" border="0" bgcolor="#666699" cellpadding="0" cellspacing="0" align="center">
    <tr> 
      <td> <table width="100%" align="center" cellspacing="1" cellpadding="4">
          <tr align="center"> 
            <td colspan="2"><strong><font color="#FFFFFF">论坛添加 </font></strong> 
              <input name="czlb" type="hidden" id="czlb" value="<%=czlb%>"> 
            </td>
          </tr>
          <tr> 
            <td width="18%" bgcolor="#FFFFFF">论坛名称</td>
            <td width="82%" bgcolor="#FFFFFF"> <input name="ltlb" type="text" id="ltlb" size="30"> 
            </td>
          </tr>
          <tr> 
            <td width="18%" bgcolor="#FFFFFF">所属类别</td>
            <td bgcolor="#FFFFFF"><%call ltlbid(czlb)%></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">论坛序号(数字)</td>
            <td bgcolor="#FFFFFF"><input name="b_order" type="text" id="b_order">
              --用于排序,从大到小 </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">论坛图标</td>
            <td width="82%" bgcolor="#FFFFFF"><input name="face" type="text" id="face" size="30"></td>
          </tr>
          <tr> 
            <td width="18%" bgcolor="#FFFFFF">论坛说明</td>
            <td width="82%" bgcolor="#FFFFFF"> <textarea name="ltsm" cols="30" rows="4" id="ltsm"></textarea> 
            </td>
          </tr>
          <tr> 
            <td width="18%" bgcolor="#FFFFFF">&nbsp;</td>
            <td bgcolor="#FFFFFF"><input type="submit" name="form1" id="form1" value="提交"> 
              <input type="reset" name="Submit2" value="重置"></td>
          </tr>
        </table></td>
    </tr>
  </table>
<%
case ("editzlb")
set rs=conn.execute("select * from ltlb where id="&request.QueryString("id"))
%>
  <table width="59%" border="0" bgcolor="#666699" cellpadding="0" cellspacing="0" align="center">
    <tr> 
      <td> <table width="100%" align="center" cellspacing="1" cellpadding="4">
          <tr align="center"> 
            <td colspan="2"><strong><font color="#FFFFFF">论坛修改</font></strong> 
              <input name="czlb" type="hidden" id="czlb4" value="<%=czlb%>"> 
              <input name="id" type="hidden" id="id" value="<%=request.QueryString("id")%>"> 
            </td>
          </tr>
          <tr> 
            <td width="19%" bgcolor="#FFFFFF">论坛名称</td>
            <td width="81%" bgcolor="#FFFFFF"> <input name="ltlb" type="text" id="ltlb" value="<%=rs("ltlb")%>" size="30"> 
            </td>
          </tr>
          <tr> 
            <td width="19%" bgcolor="#FFFFFF">论坛状态<br>
            </td>
            <td bgcolor="#FFFFFF"> <%
	  if rs("ifoff")=0 then 
	  response.Write("<a href=admin_cz.asp?czlb=offzlt&id="&rs("id")&" title=点击关闭论坛 onclick=""{if(confirm('你确实要关闭论坛吗，确定吗?')){return true;}return false;}""> <b>开放</b></a>")
	  else
	  response.Write("<a href=admin_cz.asp?czlb=offzlt&id="&rs("id")&" title=点击开放论坛 onclick=""{if(confirm('你确实要开放论坛吗，确定吗?')){return true;}return false;}""><b>关闭</b>---<font color=red>警告论坛处于关闭状态</font></a>")
	  end if
	  %>
              ( 关闭后非管理不可进入)</td>
          </tr>
          <tr> 
            <td width="19%" height="27" bgcolor="#FFFFFF">所属类别</td>
            <td bgcolor="#FFFFFF"> <%call ltlbid(cint(rs("atlb")))%> </td>
          </tr>
          <tr> 
            <td width="19%" bgcolor="#FFFFFF">论坛序号(数字)</td>
            <td bgcolor="#FFFFFF"><input name="b_order" type="text" id="b_order" value="<%=rs("b_order")%>">
              --用于排序,从大到小 </td>
          </tr>
          <tr> 
            <td width="19%" bgcolor="#FFFFFF">论坛图标</td>
            <td width="81%" bgcolor="#FFFFFF"> <input name="face" type="text" id="face" value="<%=rs("face")%>" size="30"> 
            </td>
          </tr>
          <tr> 
            <td width="19%" bgcolor="#FFFFFF">论坛说明</td>
            <td width="81%" bgcolor="#FFFFFF"> <textarea name="ltsm" cols="30" rows="4" id="ltsm"><%=rs("ltsm")%></textarea> 
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">论坛管理员</td>
            <td bgcolor="#FFFFFF"><div id="ltbz"><%=rs("ltbz")%></div></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">增删管理员</td>
            <td height="28" bgcolor="#FFFFFF"><iframe name=ad frameborder=0 width=400 height=28 scrolling=no src=adadmin.asp?id=<%=rs("id")%>></iframe></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">注册方可览帖
			
	       <%
			dim seting
			seting=split(rs("ltconfig"),"|")
			%>
			</td>
            <td bgcolor="#FFFFFF">是
              <input type="radio" name="set1" value="1" <%if cint(seting(0))=1 then response.Write("checked") end if%>>
              否 
              <input name="set1" type="radio" value="0" <%if cint(seting(0))=0 then response.Write("checked") end if%>></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">本论不许发帖</td>
            <td bgcolor="#FFFFFF">是 
              <input type="radio" name="set2" value="1" <%if cint(seting(1))=1 then response.Write("checked") end if%>>
              否 
              <input name="set2" type="radio" value="0" <%if cint(seting(1))=0 then response.Write("checked") end if%>></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">HTML代码解析</td>
            <td bgcolor="#FFFFFF">是 
              <input type="radio" name="set3" value="1" <%if cint(seting(2))=1 then response.Write("checked") end if%>>
              否 
              <input name="set3" type="radio" value="0" <%if cint(seting(2))=0 then response.Write("checked") end if%>></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">UBB代码解析</td>
            <td bgcolor="#FFFFFF">是 
              <input type="radio" name="set4" value="1" <%if cint(seting(3))=1 then response.Write("checked") end if%>>
              否 
              <input name="set4" type="radio" value="0" <%if cint(seting(3))=0 then response.Write("checked") end if%>></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td bgcolor="#FFFFFF"><input type="submit" name="form1" id="form1" value="提交"> 
              <input type="reset" name="Submit22" value="重置"></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <%end select%>
  </p>
</form>
<%
sub listboard(zid)
dim rs,sqltext,x
set rs=server.createobject("adodb.recordset")
sqltext="select ltlb,b_order,id,ifoff from ltlb where atlb='"&zid&"' order by b_order desc,id desc"
rs.open sqltext,conn,1,1
if not rs.eof then
response.Write("<ol>")   
for x=1 to rs.recordcount
response.Write("<li>")   
response.Write(rs("ltlb")&"&nbsp; &nbsp;<a href=admin_cz.asp?id="&rs("id")&"&czlb=delzlt title=删除此类别!(请不要轻易删除类别) onClick=""{if (confirm('确实要删除此类别信息吗？')){return true;}return false;}"">删除</a> ")
response.Write(" &nbsp;&nbsp;<a href=admin_index.asp?czlb=editzlb&id="&rs("id")&">修改</a>")
if rs("ifoff")=1 then 
response.Write(" -- <font color=red>关闭中</font>")
end if
response.Write("</li>" )
	rs.movenext
	next
response.Write("</ol>")
		end if
end sub
sub ltlbid(selectlt)
dim rs,sqltext,outHtm,isselect
set rs=server.createobject("adodb.recordset")
sqltext="select ltname,id from ltname order by b_order desc"
rs.open sqltext,conn,1,1
if not rs.eof then 

outHtm="<select name=""ltlbid"" id=""ltlbid"">"
do while not rs.eof
if selectlt=rs("id") then
isselect=" selected"
else
isselect=""
end if
outHtm=outHtm&"<option value="&rs("id")&isselect&">"&rs("ltname")&"</option>"
rs.movenext
loop
outHtm=outHtm&"</select>"
response.Write(outHtm)
end if
end sub
%>
</body>
</html>
