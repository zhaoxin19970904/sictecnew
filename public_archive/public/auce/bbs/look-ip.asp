<!--#include file="CONN.ASP" -->
<!--#include file="mymem.asp" -->
<!--#include file="ubbcode.asp" -->
<%
'IP换地址函数
function address(sip)
	dim str1,str2,str3,str4,sql
	dim num
	dim country,city
	dim irs
	dim conna,connstra
connstra="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
Set conna = Server.CreateObject("ADODB.Connection")
conna.Open connstra
	if isnumeric(left(sip,2)) then
	if sip="127.0.0.1" then sip="192.168.0.1"
	str1=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str2=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str3=left(sip,instr(sip,".")-1)
	str4=mid(sip,instr(sip,".")+1)
	if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then

	else
		num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
		sql="select Top 1 country,city from address where ip1 <="&num&" and ip2 >="&num&""
		set irs=server.createobject("adodb.recordset")
		irs.open sql,conna,1,1
		if irs.eof and irs.bof then 
		country="亚洲"
		city=""
		else
		country=irs(0)
		city=irs(1)
		end if
		irs.close
		set irs=nothing
	end if
	address=country&city
	else
	address="未知"
	end if
end function
%>
<table width="898" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
  <tr bgcolor="#92b9fb" background="backimg/bg1.gif"> 
    <td height="27" colspan="2" background="backimg/bg1.gif"><font color="#FFFFFF"><strong>查看用户IP来源 
      </strong></font><strong><a href="look-ip.asp">(IP搜索来源)</a></strong></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="70" colspan="2">
<%
sub checkifmaster()
if not master then 
errmess=errmess&"<li><b>此页面为管理页面.</b><li>原因:<li>你可能不具备进入此页面的权限."
call founderror(errmess)
end if
end sub
dim id,czlb
czlb=request.QueryString("czlb")
id=request.QueryString("id")
select case czlb
case "lookip"
call checkifmaster()
call listuseripcome()
case "userip"
call checkifmaster()
call seeuserip(request.QueryString("user"))
case else
call ipseeaddress()
end select 

sub listuseripcome()
dim rs,rid,thisuser
rid=request.QueryString("rid")
if rid="" then
set rs=conn.execute("select [user],[ip],[time] from borecorder where id="&id)
else
set rs=conn.execute("select user,ip,time from rely where id="&rid)
end if
thisuser=rs("user")
response.Write("<li>发帖人:"&thisuser&" 发帖时间:"&rs("time")&"<li>用户发帖IP:"&rs("ip"))
call seeuserip(thisuser)
end sub

sub seeuserip(user)
dim rs
set rs=conn.execute("select user,ip,where,hdtime,time from online where user='"&user&"'")
response.Write("<li><b>"&user&" 用户来源:</b>" )
if( not rs.eof ) or (not rs.bof) then
response.Write("<li><b>用户当前在线:</b><li>上线时间:"&rs("time")&" 活动时间:"&rs("hdtime")&"<li>在线位置:"&rs("where")&"<li> 现在IP:"&rs("ip")&" IP来源:"&address(rs("ip")))
else
set rs=conn.execute("select user,lastlogIP,lastlogin from userinfo where user='"&user&"'")
response.Write("<li><b>用户不在线:</b><li>上次登录时间:"&rs("lastlogin")&"<li>最后登录IP :"&rs("lastlogIP")&" IP来源: "&address(rs("lastlogIP")))
end if
end sub
if request.Form("submit")<>"" then
response.Write("<li>IP来源: "&address(request.Form("IP")))
end if
sub ipseeaddress()
%>
      <br>
      <br>
      <form name="form1" method="post" action="look-ip.asp">
       查看IP来源: <input name="IP" type="text" id="IP">
        <input type="submit" name="Submit" value="提交">
      </form>
<%end sub%> </td>
  </tr>
</table>
<!--#include file="end.asp" -->
</body>
</html>
