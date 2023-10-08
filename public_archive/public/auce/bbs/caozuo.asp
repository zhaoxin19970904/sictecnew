<!--#include file="conn.asp"-->
<%
if not master then 
response.cookies("rewin")("errmess")="<li><b>此页面为管理页面.</b><li>原因:<li>你可能不具备进入此页面的权限."
response.Redirect("error.asp?founderr=mess")
end if

if len(request.form("FORM1")) then
call intoadmincz()
end if

sub intoadmincz()
'on error resume next
dim czyy,czname,czid,atlt,bczuser,lb,rid,czlb,rs
czyy=request.form("czyy")
czname=request.Form("czname")
czid=request.form("czid")
atlt=request.form("atlt")
bczuser=request.form("bczuser")
lb=request.form("lb")
rid=request.Form("rid")
czlb=request.form("czlb")
if  czyy="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('请填写好 操作原因！');history.back()</script>"
Response.end
end if
dim czzl,fenl,i,fenz,nex,fen,adddatesm
czzl=array("置顶","删除","警告","精华","选登","封帖","解封","锁定","解锁","移动","封回","删回","解回","封ID","解ID")
fenl=array(0,-10,-50,100,300,0,0,0,0,0,0,0,0,0,0)
for i=0 to ubound(czzl)
if czzl(i)=czlb then
fenz=fenl(i)
end if
next
fen=fenz&fen
fen=czyy&":  分值操作: "&fen
errmess="<li>"
adddatesm=(czname&" -- "&fen)
nex="<br>&nbsp;& -- "&czname&"<li><li><a href=list.asp?lb="&lb&" >返回帖子列表</a><li><li><a href=type.asp?id="&czid&">返回帖子显示</a>"
conn.execute("insert into bbs_log([czid],[lbname],[lb],[rid],[ip],[bczuser],[czuser],[czlb],[czlr],[time]) values ('"&czid&"','"&atlt&"','"&lb&"','"&rid&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&bczuser&"','"&loginuser&"','"&czlb&"','"&adddatesm&"','"&now()&"')")
select case czlb
case "置顶"
set rs=conn.execute("select [top] from borecorder order by top desc")
conn.execute("update borecorder set [top]="&(rs("top")+1)&" where id="&czid)
errmess=errmess&czlb&rs("top")&" 操作成功! 分值操作为: "&fen&nex
case "去顶"
conn.execute("update borecorder set [top]=0 where id="&czid)
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "删除"  
conn.execute("delete from borecorder where id="&czid)
conn.execute("update userinfo set deltz=deltz+1,userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "奖惩"
conn.execute("update userinfo set userjf=userjf+"&fen&" where user='"&bczuser&"'")
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "警告"
conn.execute("update borecorder set jh=911 where id="&czid)
conn.execute("update user set userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "精华"
conn.execute("update borecorder set jh=8 where id="&czid)
conn.execute("update userinfo set jhtz=jhtz+1,userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "去精"
conn.execute("update borecorder set jh=0 where id="&czid)
conn.execute("update userinfo set jhtz=jhtz+1,userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "锁定"
conn.execute("update borecorder set lock='locked' where id="&czid)
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "解锁"
conn.execute("update borecorder set lock='' where id="&czid)
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "移帖"
conn.execute("update borecorder set lb='"&lb&"' where id="&czid)
errmess=errmess&czlb&" 到 "&zlbltname(lb)&" 操作成功! 分值操作为: "&fen&nex
case "封帖"
conn.execute("update borecorder set kill=1 where id="&czid)
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "解封"
conn.execute("update borecorder set kill=0 where id="&czid)
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex
case "封ID"
conn.execute("update userinfo set kill=1 where id="&czid)
errmess=errmess&czlb&" 操作成功! 分值操作为: 0"
errmess=errmess&"<br>&nbsp;<p><a href=listalluser.asp>返回</a> &nbsp;</p>"
case "解ID"
conn.execute("update userinfo set kill=0 where id="&czid)
errmess=errmess&czlb&" 操作成功!分传传值操作为: 0"
errmess=errmess&"<br>&nbsp;<p><a href=listalluser.asp>返回</a> &nbsp;</p>"
case "封回"
conn.execute("update rely set kill=1 where id="&rid)
errmess=errmess&czlb&" 操作成功!分传传值操作为: 0"
errmess=errmess&"<br>&nbsp;<p><a href=type.asp?id="&czid&">返回</a> &nbsp;</p>"
case"解回"
conn.execute("update rely set kill=0 where id="&rid)
errmess=errmess&czlb&" 操作成功!分值操作为: "&fen&nex
errmess=errmess&"<br>&nbsp;<p><a href=type.asp?id="&czid&">返回</a> &nbsp;</p>"
case "删回"
conn.execute("delete from rely where id="&rid)
conn.execute("update userinfo set deltz=deltz+1,userjf=userjf+"&fenz&" where user='"&bczuser&"'")
errmess=errmess&czlb&" 操作成功! 分值操作为: "&fen&nex&" <a href=type.asp?id="&czid&">返回显示帖子</a>"
case "删告"
end select
end sub
%>
<!--#include file="mymem.asp" -->

  
<table width="898" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
  <tr> 
      <td height="31" background="backimg/bg1.gif"> <strong><font color="#FFFFFF">高级管理员操作</font></strong></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF">
<%
if errmess<>"" then 
response.write(errmess)
else
call loadingmess()
end if
sub loadingmess()
dim czlb,id,rs,sql,bczuser,lb,czname,atlt,rsr
czlb=request.querystring("czlb")
id=request.querystring("id")
lb=request.QueryString("lb")
if id<>"" and lb<>"" then
set rs=conn.execute("select user,id,lb,name,kill,lock,top,jh from borecorder where id="&id)
if rs.eof or rs.bof then
response.Write("没有找到此帖子")
response.End()
end if 
bczuser=rs("user")
atlt=zlbltname(rs("lb"))
end if

select case czlb
case "置顶"
if rs("top")>0 then
czlb="去顶"
czname="去除对帖子的置顶"
else
czname="将帖子固定于论坛顶部"
end if
case "删除"
czname="删除帖子"
case "移帖"
czname="移动帖子"
case "封帖"
if rs("kill")=0 then
czlb="封帖"
czname="封闭帖子"
else
czname="解除封闭"
czlb="解封"
end if
case "锁定"
if rs("lock")="locked" then
czlb="解锁"
czname="解除锁定"
else
czlb="锁定"
czname="锁定帖子"
end if
case "奖励"
czname="奖励给用户积分"
case "精华"
if rs("jh")=8 then
czlb="去精"
czname="去除此帖在精华区"
else
czname="将帖置为精华"
end if
case "封ID"
   set rs=conn.execute("select user,kill from userinfo where id="&id)
   if rs("kill")=1 then 
   czlb="解ID"
   czname="解除对用户ID的封闭"
   else
   czname="封闭用户ID"
   end if
   bczuser=rs("user")
case "封回"
set rsr=conn.execute("select user,kill,lb from rely where id="&request.QueryString("rid"))
if rsr.bof or rsr.eof then 
response.Write("无此帖子,操作失误")
response.End()
end if
bczuser=rsr("user")
if rsr("kill")=0 then
czlb="封回"
czname="锁定这个回复"
else
czname="解除对这个回复锁定"
czlb="解回"
end if
case "删回"
set rsr=conn.execute("select user,kill,lb from rely where id="&request.QueryString("rid"))
czname="删除回复的帖子"
bczuser=rsr("user")
end select
%>
<form name="FORM1" method="post" action="caozuo.asp?lb=<%=lb%>">
	  <table width="100%" border="0">
          <tr> 
            <td><strong>操作内容: <font color="#FF0000"><%=czname%></font></strong></td>
          </tr>
          <tr> 
            <td><input type="hidden" name="bczuser" value="<%=bczuser%>"> 
              <input name="atlt" type="hidden" id="atlt" value="<%=atlt%>"> 
              <input type="hidden" name="czname" value="<%=czname%>"> 
              <input type="hidden" name="czid" value="<%=id%>"> 
              <input name="czlb" type="hidden" id="czlb" value="<%=czlb%>"> 
              <input name="lb" type="hidden" id="lb" value="<%=lb%>"> 
              <input name="rid" type="hidden" id="rid" value="<%=request.QueryString("rid")%>">
              操作原因: 
              <select name="kczyy" size=1 onChange="document.FORM1.czyy.value=kczyy.value">
                <option value="">自定义</option>
                <option value="灌水">灌水</option>
                <option value="广告">广告</option>
                <option value="奖励">奖励</option>
                <option value="好文章">好文章</option>
                <option value="内容不符">内容不符</option>
                <option value="重复发帖">重复发帖</option>
                <option value="被杂志选登">被杂志选登</option>
                <option value="选为精华">选为精华</option>
              </select> <input name="czyy" type="text" id="czyy" size="60" ></td>
          </tr>
          <tr> 
            <td> <% if request.querystring("czlb")="奖惩" then%>
              分值: 
              <select name=fen>
                <option value="0" selected>0</option>
                <script language=JavaScript><!--               
          for(i=-2000;i<2001;i++) document.write('<option>'+i)               
            //-->
         </script>
              </select> <%end if%> <%
if request.querystring("czlb")="移帖" then
call ltnamelist()
end if
%> </td>
          </tr>
          <tr> 
            <td align="center"> <input name=FORM1 type=submit  value="提 交"> &nbsp; 
              <input name="Submit" type="button" onClick="history.go(-1);" value="返 回"> 
            </td>
          </tr>
        </table>
		</form>
		<%end sub%>
		</td>
    </tr>
  </table>

<%
sub ltnamelist()
dim rs
set rs=conn.execute("select ltlb,id from ltlb order by b_order desc")
response.write" &nbsp; 请选择要移到的论坛&nbsp; "
response.write"<select name=""ltlb"">"
if not rs.eof then
do while not rs.eof
response.write"<option value="""&rs("id")&""">"&rs("ltlb")&"</option>" 
rs.movenext
loop
response.write"</select>"
end if
end sub
%>
<!--#include file="end.asp" -->
</body>
</html>
