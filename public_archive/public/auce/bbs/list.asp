<!--#include file="conn.asp"-->
<!--#include file="ubbcode.asp" -->
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
dim lb,rs,sql
lb=request.querystring("lb")
%>
<script language="JavaScript">
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
//	if (targetImg.src.indexOf("nofollow")!=-1){return false;}
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			targetImg.src="pic/nofollow.gif";
			if (targetImg.loaded=="no"){
           var thefollowtable="followTd"+t_id
           document.all[thefollowtable].innerHTML="<IFRAME name=main src=loadtree.asp?id="+t_id+"&rlyid="+b_id+" frameBorder=0 width=100% height=180></IFRAME>";
//        document.frames["hiddenframe"].location.replace("loadtree.asp?id="+t_id+"&rlyid="+b_id);
			}
		}else{
			targetDiv.style.display="none";
			targetImg.src="pic/follow.gif";
		}
	}
}
</script>
<!--#include file="mymem.asp" -->

<table width="898" border=0 align="center" cellpadding=0 cellspacing=0>
  <tbody>
    <tr>
	<tr> 
      <td bgcolor="#FFFFFF" >&nbsp; </td>
    </tr>
      <td width="787" bgcolor="#FFFFFF" >

	<table width="100%" align="center" bgcolor="#92b9fb" border="0" cellpadding="0" cellspacing="1">
          <tr> 
            <td height="27" background="backimg/bg1.gif"> &nbsp;
              <IFRAME name="hiddenframex" frameBorder=0 width=0 height=0></IFRAME>
               <%call onlineusersum("aall")%>
              <b><font color="#FFFFFF"> <a href=loadtree.asp?listonline=user target="hiddenframex" onClick="document.all.listonline.innerHTML='正在读取用户列表，请稍侯……'">[显示详细列表]</a></font></b></td>
          </tr>
          <tr> 
            
          <td bgcolor="#FFFFFF" > 
            <div id="listonline"></div></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" >&nbsp; </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" ><a href="AddIn.asp?lb=<%=lb%>&czlb=addtz"><img src="pic/postnew.gif" alt="发新帖" width="72" height="21" border="0"></a>&nbsp;&nbsp; 
        <a href="tpaddIn.asp?lb=<%=lb%>"><img src="PIC/votenew.gif" alt="发起投票" width="72" height="21" border="0"></a> 
        <%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from borecorder where lb='"&lb&"'  order by top desc,retime desc,id desc" 
rs.open sql,conn,1,1

dim MaxPerPage
MaxPerPage=20

'取得页数,并判断用户输入的是否数字类型的数据，如不是将以第一页显示
dim text,checkpage,i,CurrentPage
text="0123456789"
 Rs.PageSize=MaxPerPage
for i=1 to len(request("page"))
   checkpage=instr(1,text,mid(request("page"),i,1))
   if checkpage=0 then
      exit for 
   end if
next

If checkpage<>0 then
      If NOT IsEmpty(request("page")) Then
        CurrentPage=Cint(request("page"))
        If CurrentPage < 1 Then CurrentPage = 1
        If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
      Else
        CurrentPage= 1
      End If
      If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
   CurrentPage=1
End if

call tabletop(true,"top")
if (not rs.eof) and (not rs.bof) then
call list
call showpages("&lb="&lb)
else
response.Write("<table width=770 align=center bgcolor=#92b9fb border=0 cellpadding=4 cellspacing=1><tr><td bgcolor=#FFFFFF  height=25>论坛暂无任何主题</td></tr></table>")
end if
%>
      </td>
    </tr>
    <tr> 
      <td>
<%
'显示帖子的子程序
Sub list()
%>
 <table width="100%" align="center" height="64" bgcolor="#92b9fb" border="0" cellpadding="4" cellspacing="1">
          <tr align="center" bgcolor="#0053A6"> 
            <td height="27" colspan="2" background="backimg/bg1.gif"><font color="#FFFFFF">状态 
              </font></td>
            <td width="379"  height="27"  background="backimg/bg1.gif"><font color="#FFFFFF">主题 
              </font></td>
            <td width="28"  height="27" background="backimg/bg1.gif"  ><font color="#FFFFFF">新窗</font></td>
            <td width="55" height="27" background="backimg/bg1.gif" ><font color="#FFFFFF"> 
              人气/回复</font></td>
            <td width="82" height="27" background="backimg/bg1.gif"><font color="#FFFFFF"> 
              作者 </font></font></td>
            <td width="122" height="27" background="backimg/bg1.gif"><font color="#FFFFFF">发帖时间</font></td>
          </tr>
          <%
dim i,crely,top,x
i=1
  do while not rs.eof 
  crely=rs("rely")
  top=rs("top")
%>
          <tr> 
            <td width="20" height="19" align="center" bgcolor="#EEEEEE"> <% if top>0 then 
    response.write"<img src=""pic/istop.gif"" alt=""此贴被置顶过的贴子"">"
    elseif rs("lock")="locked" then
    response.write"<img src=""pic/locked.gif"" alt=""此贴子已被锁定"">"
    elseif rs("istp")<>0 then
    response.write"<img src=""pic/tp.gif"" alt=""此帖为投票帖子"">" 
    elseif rs("jh")=8 then 
    response.write"<img src=""pic/jhinfo.gif"" alt=""此贴为精华贴子"">" 
    elseif rs("jh")=911 then
    response.write"<img src=""pic/jg.gif"" alt=""此贴是被警告的贴子!"">" 
    elseif rs("jh")=888 then
    response.write"<img src=""pic/book.gif"" alt=""此贴被杂志选登!"">" 
	else
	response.write"<img src=""pic/folder.gif"" alt=""普通开放贴子"">"
end if

%> </td>
            <td width="20" height="19" bgcolor="#EEEEEE" ><img src="<%=rs("heat")%>" alt="发贴心情" width="20" height="20"></td>
            <td width="379" onmouseover=javascript:this.bgColor='#F9FBFD' onmouseout=javascript:this.bgColor='#FFFFFF' bgcolor="#FFFFFF" height="19"  > 
              <% 
dim ts
ts=("发贴人:"&rs("user")&"<br>发贴时间:"&rs("time")&"<br>最后回复:"&rs("lastly")&"<br>最后回复时间:"&rs("retime"))
    if crely>0 then 
	response.write "<img loaded=""no"" src=""pic/follow.gif"" id=""followImg"& rs("id") &"""style=""cursor:hand;"" onclick=""loadThreadFollow("& rs("id")&","&lb&")"" title=展开贴子列表>"
    else
	response.write "<img src=""pic/nofollow.gif"" id=""followImg"& rs("id") &""">"
    end if
response.write"<a href=type.asp?id="&rs("id")&" title="""&ts&""" >"&dvHTMLEncode(rs("name"))&"</a>" 
        '当回复达到多页时显示
		if crely>10 then
		dim disppagesd
		response.Write("[<b><img src=backimg/multipage.gif>")
		if CInt(crely/10+0.5)>8 then 
		disppagesd=8
		else
		disppagesd=CInt(crely/10+0.5)
		end if
        for x=1 to disppagesd
        Response.write" <a href='type.asp?page="&x&"&lb="&lb&"&id="&rs("id")&"'><font color=red>"&x&"</font></a>"
        next
		if disppagesd=8 then 
		response.Write(" ... <a href='type.asp?page="&CInt(crely/10+0.5)&"&lb="&lb&"&id="&rs("id")&"'><font color=red>"&CInt(crely/10+0.5)&"</font></a>")
		end if
		response.Write("</b>]")
        end if
%> </td>
            <td width="28" height="19" align="center"  bgcolor="#EEEEEE"  > <a href="type.asp?id=<%=rs("id")%>" target="game"><img src="PIC/window.gif" width="20" height="20" border="0" alt="在新窗口打开贴子"></a></td>
            <td width="55" height="19" align="center" bgcolor="#EEEEEE" > <%=rs("click")&"/"&crely%></td>
            <td width="82" height="19" align="center" bgcolor="#ffffff"> <a href="usertype.asp?user=<%=rs("user")%>" title="点击查看会员资料" ><%=rs("user")%></a></td>
            <td width="122" height="19" bgcolor="#EEEEEE"><%=rs("time")%></td>
          </tr>
          <%
   response.write "<tr style=""display:none"" id=""follow"& rs("id") &"""><td colspan=8 id=""followTd"& rs("id") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& rs("id") &","&lb&")"">正在读取关于本主题的跟贴，请稍侯……</div></td></tr>"
   if i >= MaxPerpage then exit do
   rs.movenext
   i=i+1
   loop
%>
        </table>
        <%
End sub
rs.close
%> </td>
    </tr>
</table>
<%
'标准分页子程序URLparameter为URL所带数
sub showpages(URLparameter)
dim pages,soonhost,pagename,nx,x,ISselect,tenpages
pages=cint(request.QueryString("page"))
if pages=0 then
pages=1
end if
if pages>rs.pagecount then
pages=rs.pagecount
end if
pagename=request.ServerVariables("SCRIPT_NAME")

response.Write("<table width=770 border=0 align=center cellpadding=2 cellspacing=1 bgcolor=#92b9fb>")
response.Write(" <tr bgcolor=#FFFFFF>")
response.Write("  <td width=734 class=main1>")

tenpages=cint(pages/10-0.5001)
nx=tenpages*10
response.Write("页次:<b>"&pages&"/"&rs.pagecount&"</b>&nbsp;每页<b>"&MaxPerPage&"</b>&nbsp;主题数<b>"&rs.recordcount&"</b>&nbsp;")
if request("page")>1 then
response.write "&nbsp;<a href='"&pagename&"?page=1"&URLparameter&"' title=首页><font face=webdings>9</font></a>"
else
response.write "&nbsp;<font face=webdings>9</font>"
end if

if nx>=10 then
response.write "&nbsp;<a href='"&pagename&"?page="&(nx-10)&URLparameter&"' title=上十页><font face=webdings>7</font></a>"
else
response.write "&nbsp;<font face=webdings>7</font>"
end if

for x=1 to 10
nx=nx+1
if pages=nx then
response.write ("<b>&nbsp;<font color=red>"&nx&"</font></b>")
else
Response.write "<b>&nbsp;<a href='"&pagename&"?page="&nx&URLparameter&"'>"&nx&"</a></b>"
end if
if nx>=rs.pagecount then exit for
next

if rs.pagecount>=10 and nx<rs.pagecount then
response.write "&nbsp;<a href='"&pagename&"?page="&(nx+1)&URLparameter&"' title=下十页><font face=webdings>8</font></a>"
else
response.write "&nbsp;<font face=webdings>8</font>"
end if
if pages<rs.pagecount then
response.write "&nbsp;<a href='"&pagename&"?page="&rs.pagecount&URLparameter&"' title=尾页><font face=webdings>:</font></a>"
else
response.write "&nbsp;<font face=webdings>:</font>"
end if

response.Write("</td><td width=""25"" class=main1 >")

response.Write("<select name=pagd onChange=setTimeout(location.href=this.value,3000)>")
soonhost=1
do while not soonhost = (rs.PageCount+1) 
if (soonhost)=pages then
ISselect=" selected"
else
ISselect=""
end if
response.Write("<option value="&pagename&"?Page="&soonhost&URLparameter&ISselect&">"&soonhost&"</option>")
soonhost=soonhost+1
loop

response.Write("</select> </td></tr></table>")
end sub
call tabletop(true,"down")
%>
<TABLE width="898" border=0 align="center" cellPadding=3 cellSpacing=1 bgcolor="#FFFFFF">
  <TBODY>
    <TR> 
      <TD height="27"  align="right" > <form name="form1" method="get" action="search.asp">
          <select name="shlb" id="shlb">
            <option value="name">按发帖人</option>
            <option value="lyname">帖子题目</option>
            <OPTION value=1>前1天的</OPTION>
            <OPTION value=2>前2天的</OPTION>
            <OPTION  value=5>前5天的</OPTION>
            <OPTION value=10>前10天的</OPTION>
            <OPTION value=20>前20天的</OPTION>
            <OPTION value=30>前30天的</OPTION>
            <OPTION value=45>前45天的</OPTION>
            <OPTION  value=60>前60天的</OPTION>
            <OPTION value=75>前75天的</OPTION>
            <OPTION  value=365>前1年的</OPTION>
            <OPTION value=0>所有的帖子</OPTION>
          </SELECT>
          <input name="shname" type="text" id="shname" size="20">
          <input type="submit" name="Submit" value="搜索">
          <%call listchangeLt()%>
         </form></TD>
    </TR>
    <TR bgColor=#a9d46d> 
      <TD height="27" background="backimg/bg1.gif" bgcolor="#a9d46d"><font color=#FFFFFF><b> 
        论坛图例</b></font></TD>
    </TR>
    <TR> 
      <TD bgColor=#f4faed colSpan=2> <TABLE cellSpacing=4 cellPadding=0 width="92%" align=center 
border=0>
          <TBODY>
            <TR> 
              <TD><IMG src="PIC/folder.gif" width="16" height="16"> 开放主题</TD>
              <TD><img src="PIC/jg.gif" width="16" height="16"> <a href="search.asp?jh=911">被警告帖子</a></TD>
              <TD><IMG src="PIC/locked.gif" width="16" height="16"> 锁定主题</TD>
              <TD><IMG src="pic/istop.gif" width="15" height="17"> 固定主题 </TD>
              <TD><IMG src="PIC/jhinfo.gif" width="15" height="16"> <a href="search.asp?jh=8">精华帖子 
                </a></TD>
              <TD><IMG src="PIC/tp.gif" width="13" height="15"> 投票帖子 </TD>
            </TR>
            <TR> 
              <TD width="100%" 
      colSpan=6></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
<!--#include file="end.asp" -->
</body>
</html>





















