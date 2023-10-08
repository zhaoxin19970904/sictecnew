<!--#include file="CONN.ASP" -->
<%
dim lb
if not userislogin then
founderr=true
errmess=errmess&error1
end if
if userloginlock=1 then
founderr=true
errmess=errmess&error2
end if
lb=request.QueryString("lb")
if len(request.form("FORM1")) then
call addtpdate()
end if
sub addtpdate()
if founderr then
exit sub
end if
dim rs,sqltext,tply,ltsel,x,i,bz,bzc,from_address,tpid,id
x=0
tply=replace(trim(request.Form("tply")),vbCrlf,"|")
		ltsel=split(tply,"|")
		for i=0 to ubound(ltsel)
        if ltsel(i)<>"" then
        x=x+1		
        bz=ltsel(i)+"|"+bz
        bzc="0|"+bzc
        end if
        next
if x<2 or x>10 then
founderr=true
errmess=errmess&"<li>产生错误:<li>原因:投票项目不能少于2项和大于10项的!"
exit sub
end if 
if request.Form("name")="" or request.Form("ly")="" then
founderr=true
errmess=errmess&"<li>错误:主题和主题内容都填项,请注意填定完整"
exit sub
end if
call updateftuser(true)
set rs=server.createobject("adodb.recordset")
rs.open "lttp",conn,1,3
rs.Addnew
rs("tply")=bz
rs("count")=bzc
rs("name")=request.form("name")
rs("datetime")=now()
rs("tplb")=request.form("tplb")
rs("timeout")=request.form("timeout")
tpid=rs("id")
rs.update

set rs=server.createobject("adodb.recordset")
rs.open "borecorder",conn,1,3

'添加一条记录到数据库
rs.addnew
rs("istp")=tpid
rs("name")=request.form("name")
rs("heat")=request.form("heat")
rs("ly")=request.form("ly")
rs("lb")=request.form("lb")
rs("user")=loginuser
rs("time")=now()
rs("retime")=now()
rs("ip")=Request.ServerVariables("REMOTE_ADDR")
id=rs("id")
rs.update
founderr=true
errmess=errmess&"<li>你的帖子已经成功的发到论坛了,3 秒钟自动返回你所发表的帖子."
errmess=errmess&"<script>window.tm = setInterval(""location.href='type.asp?id="&id&"'"", 3000)</script>"
errmess=errmess&"<li> <li>"&thispage_name(request.ServerVariables("SCRIPT_NAME"),"url")
errmess=errmess&"<li> <li><a href=type.asp?id="&id&">你所发表的帖子</a>"
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
   document.form1.name.focus();
   return false;
 }
  if(document.FORM1.tply.value.length<1)
 {
   alert("请输入投票项目.");
   document.form1.name.focus();
   return false;
 }
   if(document.FORM1.ly.value.length<1)
 {
   alert("请填写好你要发表的内容");
   document.FORM1.ly.focus();
   return false;
  }
}
//-->
</SCRIPT>
<%
if founderr=true then
call founderror(errmess)
end if
%>
<FORM action=<%=request.ServerVariables("SCRIPT_NAME")%>?lb=<%=lb%> method=post name=FORM1 id="form1"   onsubmit="return form1_onsubmit()" language=javascript>
  <table width="898" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
    <tr bgcolor="#6699CC"> 
      <td height="27" colspan="3" background="backimg/bg1.gif"> <font color="#FFFFFF"><strong>论坛发起投票</strong></font><b> 
        <input name="lb" type="hidden" id="lb" value="<%=lb%>">
        </b></td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td width="148"> <b>主题标题</b> <SELECT name=fonta onchange="document.FORM1.name.value+=fonta.value">
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
        </SELECT> </td>
      <td width="363"> <input type="text" name="name" size="60" maxlength="100" > 
      </td>
      <td width="240">有效 
        <select name="timeout" >
          <option selected>永久</option>
          <option value="10">10</option>
          <option value="15">15</option>
          <option value="30">30</option>
          <option value="60">60</option>
          <option value="90">90</option>
        </select>
        天</td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td width="148"> <p><strong>投票项目</strong><br>
          每一换行为一个项目<br>
          总项目不能少于两个或大于10个<br>
          (不能使用HTML标签和UBB标签)</p></td>
      <td colspan="2"> <textarea name="tply" cols="60" rows="6" ></textarea> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE"""">
      <td width="148" height="30"><strong>文件上传</strong> <a href="#" title="gif,jpg,bmp,jpeg,png">类型</a></td>
      <td colspan="2"><IFRAME name=open src="upfilex.asp?lb=<%=lb%>" frameBorder=0 width="610"  height="30"></IFRAME></td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td width="148" valign="top"> <strong>发帖心情:</strong></td>
      <td colspan="2" rowspan="2"> <!--#include file="getubb.asp" --> <p> 
          <textarea name="ly" rows="8" wrap="file"  class=cid onKeyDown=ctlent() cols="70"></textarea>
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
%>
        </p></td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td width="148"> <strong><a href="helpubb.asp">内容</a></strong><br>
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
%>
        <br> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"#E0E4FE""""> 
      <td colspan="3"> <div align="center"> 
          <p><font color="#0000FF"> 
            <input name=FORM1 type=submit  value="提 交">
            &nbsp;&nbsp; 
            <input type="reset" toLowerCase name="Reset" value="重 填">
            </font></p>
        </div></td>
    </tr>
  </table>
  </center>
  </div>
</FORM>
<!--#include file="end.asp" -->
</body>
</html>
