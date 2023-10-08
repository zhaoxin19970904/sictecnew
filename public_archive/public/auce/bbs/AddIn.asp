<!--#include file="CONN.ASP" -->
<%
dim rs,sql,name,ly,czlb,id,rid,lb
if not userislogin then
founderr=true
errmess=errmess&error1
end if
if userloginlock=1 then
founderr=true
errmess=errmess&error2
end if
'加入数据子程序
sub addIntoData()
dim czlb,id,rid,lysting
id=request.Form("id")
rid=request.Form("rid")
czlb=request.Form("czlb")
if id<>"" and czlb<>"addtz" then
call checkiflocked(id)
end if
if founderr then
exit sub
end if
set rs=server.createobject("adodb.recordset")
select case czlb
case "addtz"
call updateftuser(true)
sql="select * from borecorder"
rs.open sql,conn,1,3
rs.addnew
rs("retime")=now()
lysting=request.form("ly")
case "edittz"
sql="select * from borecorder where id="&id
rs.open sql,conn,1,3
lysting=request.form("ly")&vbCrlf&vbCrlf&" [color=#FF00FF]此帖已被修改过,修改时间为:[/color]"&now()
case "retz"
call updateftuser(false)
sql="select * from rely"
rs.open sql,conn,1,3
rs.addnew
rs("rid")=id
lysting=request.form("ly")
case "editretz"
sql="select * from rely where id="&rid
rs.open sql,conn,1,3
lysting=(request.form("ly")&vbCrlf&vbCrlf&" [color=#FF00FF]此帖已被修改过,修改时间为:[/color]"&now())
case "vote"
call updateftuser(false)
sql="select * from rely"
rs.open sql,conn,1,3
rs.addnew
lysting=request.form("ly")
rs("rid")=id
case "revote"
call updateftuser(false)
sql="select * from rely"
rs.open sql,conn,1,3
rs.addnew
lysting=request.form("ly")
rs("rid")=id
case "editvote"
sql="select * from rely where id="&rid
rs.open sql,conn,1,3
lysting=request.form("ly")&vbCrlf&vbCrlf&" [color=#FF00FF]此帖已被修改过,修改时间为:[/color]"&now()
end select
rs("name")=request.form("name")
rs("heat")=request.form("heat")
rs("ly")=lysting
rs("lb")=request.form("lb")
rs("user")=loginuser
rs("time")=now()
rs("ip")=Request.ServerVariables("REMOTE_ADDR")
if id="" then 
id=rs("id")
end if
rs.update
founderr=true
errmess=errmess&"<li>你的帖子已经成功的发到论坛了,3 秒钟自动返回你所发表的帖子."
errmess=errmess&"<script>window.tm = setInterval(""location.href='type.asp?id="&id&"'"", 3000)</script>"
errmess=errmess&"<li> <li>"&thispage_name(request.ServerVariables("SCRIPT_NAME"),"url")
errmess=errmess&"<li> <li><a href=type.asp?id="&id&">你所发表的帖子</a>"
end sub

if request.form("sublit")<>"" then
addIntoData()
else
czlb=request.QueryString("czlb")
lb=request.QueryString("lb")
id=request.QueryString("id")
rid=request.QueryString("rid")
if czlb="" or lb="" then
founderr=true
errmess=errmess+"<li>对不起发现参数错误,这可能是由于你没有从正确的链接进入.请回<a href=index.asp>论坛</a>"
end if
if (czlb<>"addtz") and id="" then
founderr=true
errmess=errmess+"<li>对不起!发现参数错误,这可能是由于你没有从正确的链接进入.请回<a href=index.asp>论坛</a>"
end if
if (czlb="editretz" or czlb="editvote" or czlb="revote") and rid="" then
founderr=true
errmess=errmess+"<li>对不起!发现参数错误,这可能是由于你没有从正确的链接进入.请回<a href=index.asp>论坛</a>"
end if
select case czlb
case "addtz"
name="" 
ly=""
case "edittz"
set rs=conn.execute("select name,ly from borecorder where id="&id&" and user='"&loginuser&"'")
if not rs.eof then
name=rs("name")
ly=rs("ly")
else
founderr=true
errmess=errmess+"<li>对不起!无法显示你要编辑的帖子.<li>原因: 你可能没有编辑此帖子的权限请注意从正确链接 <a href=index.asp>进入</a>"
end if

case "retz"
set rs=conn.execute("select name,lock from borecorder where id="&id)
if not rs.eof then
name="[回复:]"&rs("name")
else
founderr=true
errmess=errmess+"<li>对不起!没有找到你回复所指向的帖子.请从正确链接<a href=index.asp>进入</a>"
end if

case "editretz"
set rs=conn.execute("select name,ly from rely where id="&rid&" and user='"&loginuser&"'")
if not rs.eof then
name=rs("name")
ly=rs("ly")
else
founderr=true
errmess=errmess+"<li>对不起!无法显示你要编辑的帖子.<li>原因: 你可能没有编辑此帖子的权限.请注意从正确链接 <a href=index.asp>进入</a>"
end if

case "vote"
set rs=conn.execute("select name,user,time,ly,lock from borecorder where id="&id)
if not rs.eof then
name="[引用]:"&rs("name")
ly="[quote][color=#FF1493]以下是引用 "&rs("user")&" 在 "&rs("time")&"  的发言：[/color]"&vbCrlf&rs("ly")&"[/quote]"
else
founderr=true
errmess=errmess+"<li>对不起!没有找到你引用所指向的帖子.请从正确链接<a href=index.asp>进入</a>"
end if

case "editvote"
set rs=conn.execute("select name,ly from rely where id="&rid&" and user='"&loginuser&"'")
if not rs.eof then
name=rs("name")
ly=rs("ly")
else
founderr=true
errmess=errmess+"<li>对不起!无法显示你要编辑的帖子.<li>原因: 你可能没有编辑此帖子的权限.请注意从正确链接<a href=index.asp>进入</a>"
end if
case "revote"
set rs=conn.execute("select name,user,time,ly from rely where id="&rid)
if not rs.eof then
name="[引用]:"&rs("name")
ly="[quote][color=#FF1493]以下是引用 "&rs("user")&" 在 "&rs("time")&"  的发言：[/color]"&vbCrlf&rs("ly")&"[/quote]"
else
founderr=true
errmess=errmess+"<li>对不起!没有找到你引用所指向的帖子.请从正确链接<a href=index.asp>进入</a>"
end if
end select
end if
'检查帖子是否被锁定
sub checkiflocked(lockid)
dim rs
set rs=conn.execute("select lock from borecorder where id="&lockid)
if rs("lock")="locked" then
founderr=true
messerr=messerr&"<li>对不起!此帖子已被管理员锁定,不再接受回复.<li><a href=type.asp?id="&lockid&">返回帖子</a>"
end if
end sub
%>
<!--#include file="mymem.asp" -->
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--

function form1_onsubmit() 
{
  if(document.FORM1.name.value.length<1)
 {
   alert("你忘了输入标题了");
   document.FORM1.name.focus();
   return false;
 }
   if(document.FORM1.ly.value.length<1)
 {
   alert("请填写好你要发表的内容");
   document.FORM1.ly.focus();
   return false;
  }
}
function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="sublit"||tempobj.Reset.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
//-->
</SCRIPT>

<FORM name=FORM1 onSubmit="return form1_onsubmit()" action=AddIn.asp?lb=<%=lb%>&czlb=ok method=post>
  <table width="898" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="787" height="19"> </td>
    </tr>
    <tr>
      <td> 
<%
if founderr then
founderror(errmess)
end if
call tabletop(true,"top")
%>
        <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
          <tr bgcolor="#92b9fb" background="backimg/bg1.gif"> 
            <td height="27" colspan="2" background="backimg/bg1.gif"> <input name="lb" type="hidden" id="lb" value="<%=lb%>"> 
              <input name="czlb" type="hidden" id="czlb" value="<%=czlb%>"> 
              <input name="id" type="hidden" id="id" value="<%=id%>"> 
              <input name="rid" type="hidden" id="rid" value="<%=rid%>"> 
              <font color="#FFFFFF"><strong>论坛发帖子 </strong></font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%"><b>主题标题</b><SELECT name=fonta onchange="document.FORM1.name.value+=fonta.value">
              <OPTION selected value="">选择话题</OPTION> 
			  <OPTION value=[原创]>[原创]</OPTION> 
              <OPTION value=[转帖]>[转帖]</OPTION> 
			  <OPTION value=[灌水]>[灌水]</OPTION> 
              <OPTION value=[讨论]>[讨论]</OPTION> 
			  <OPTION value=[求助]>[求助]</OPTION> 
              <OPTION value=[推荐]>[推荐]</OPTION> 
			  <OPTION value=[公告]>[公告]</OPTION> 
              <OPTION value=[注意]>[注意]</OPTION> 
			  <OPTION value=[贴图]>[贴图]</OPTION>
              <OPTION value=[建议]>[建议]</OPTION> 
			  <OPTION value=[下载]>[下载]</OPTION>
              <OPTION value=[分享]>[分享]</OPTION>
			  </SELECT>
			  </td>
            <td width="81%"> <input name="name" type="text" id="tzname" value="<%=name%>" size="60" maxlength="100" > 
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%"><strong>文件上传</strong> <a href="#" title="gif<br>jpg<br>bmp<br>jpeg<br>png">类型</a> <br> </td>
            <td bgcolor="#FFFFFF"><IFRAME name=open src="upfilex.asp?lb=<%=lb%>" frameBorder=0 width="610"  height="30"></IFRAME></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%"> <strong>发帖心情</strong>: </td>
            <td width="81%" rowspan="2"> <!--#include file="getubb.asp" --> <textarea name="ly" rows="12" wrap="file" cols="75" title="按Ctrl+Enter可提交帖子" onkeydown=ctlent()><%=ly%></textarea> 
              <br> 
<%
call listpicimg()
sub listpicimg()
dim i
for i=1 to 28
	if len(i)=1 then i="0" & i
	response.write "<img src=""pic/em"&i&".gif"" border=0 onclick=""insertsmilie('[em"&i&"]')"" style=""CURSOR: hand"">&nbsp;"
if i=14 then
response.write"<br>"
end if
next
end sub
%> </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="19%" height="231"> <strong>内容</strong><br>
<%
if cint(bbs_seting(2))=1 then
response.Write("HTML代码:不支持")
else
response.Write("HTML代码:支持")
end if
if cint(bbs_seting(3))=0 then
response.Write("<br>UBB代码:不支持")
else
response.Write("<br><a href=helpubb.asp>UBB</a>代码:支持")
end if
%></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2" align="center"> 
              <input name=sublit type=submit id="sublit" value="提 交"> &nbsp;&nbsp; 
              <input type="reset" tolowercase name="Reset" value="重 填"> </td>
          </tr>
        </table>
        <%call tabletop(true,"down")%>
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
<script language=javascript>
ie = (document.all)? true:false
if (ie){
function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.FORM1.submit();}}
}
</script>
</FORM>
<!--#include file="end.asp" -->
</body>
</html>
