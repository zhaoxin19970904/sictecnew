<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
curclassid=trim(session("classid"))
if curclassid="" then
   response.redirect "error.asp?info=�Բ������Ѿ������ˣ������½��룡"
end if
set rs = createobject("ADODB.recordset")
set rss = createobject("ADODB.recordset")
%>
<script language="javascript">

function selectclass(curid){
		
			if (curid!=0) {
		        window.location.href="class_index.asp?selectclassid="+curid
		}
	}
function personalinfo(userid){
		
			
		window.open("popwindow_perinfo.asp?userID="+userid,"top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=580,height=400")
		
	}	
	
function modifyinfo(){
		
			
		window.open("popwindow_modifyinfo.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=580,height=600")
		
	}
function newclass(){
		
			
		window.open("../find_class1.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=770,height=500")
		
	}
function searchmate(){
		
			
		window.open("popwindow_search.html","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=570,height=280")
		
	}	
function ClassBBSList(){
		
			
		window.open("popwindow_liuyan.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=580,height=450")
		
	}
function ClassBBSList2(){
		
			
		window.open("popwindow_Perliuyan.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=570,height=550")
		
	}	
function PopWindowUpload(){
		
			
		window.open("PopWindowUpload.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=570,height=380")
		
	}
function PopWindowEmailList(){
		
			
		window.open("PopWindowEmailList.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=600,height=600")
		
	}	
function personallogout(){
               if (confirm("��ȷ��Ҫע��������༶�������"))
	{		
			
		window.location.href="class_index.asp?act=del"
	}	
	}
function manage(){
		
			
		window.open("popwindow_manage.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=570,height=500")
		
	}
function PopWindowEmailMates(){
		
			
		window.open("PopWindowEmailMates.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=580,height=450")
		
	}				
</script>	
<html>
<head>
<title>����ͬѧ¼-��ʦͨѶ¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
<meta name="web_designner" content="�����">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="5">
<table border="0" width="750" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="left" valign="top" width="204"><img src="../school_images/s_top.gif" width="207" height="72"></td>
    <td align="center" valign="middle" width="470"> <a href="http://www.qzone.com" target="_blank"> 
      <script language="JavaScript">
var f = Math.random();
f=f*100+0.5;
id=Math.round(f);
document.write("<a href=http://61.134.4.193/ad/adclick.asp?pageid=49&id="+id)
document.write("><img src=http://61.134.4.193/ad/adimg.asp?pageid=49&id="+id+" border=0></a>")
</script>
      </a></td>
    <td align="center" valign="top" class="topic"><a href="http://music.269.net" target="_blank" class=8>20000��MP3</a><br>
      <a href="http://soft.269.net" target="_blank" class=8>�������</a><br>
      <a href="http://love.269.net" target="_blank" class=8>��Ů˧��</a></td>
  </tr>
</table>
<table border="0" width="750" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="1" colspan="2"><img border="0" src="images/space.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td bgcolor="#006699" height="1" colspan="2"><img border="0" src="images/space.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td bgcolor="#CDE9F3" height="19"  colspan="2" valign="bottom"><img border="0" src="http://cx.269.net/bbs/images/go.gif" width="15" height="13"> 
      <img border="0" src="http://cx.269.net/bbs/images/home.gif" width="15" height="13"><a href="http://www.269.net"  class="5">��԰��ҳ</a> 
      <img border="0" src="http://cx.269.net/bbs/images/music.gif" width="15" height="13"><a href="http://music.269.net"  class="5">����</a> 
      <img border="0" src="http://cx.269.net/bbs/images/club.gif" width="15" height="13"><a href="http://cx.269.net" class="5">����</a> 
      <img border="0" src="http://cx.269.net/bbs/images/ebuy.gif" width="15" height="13"><a href="http://ebuy.269.net" class="5">����</a> 
      <img border="0" src="http://cx.269.net/bbs/images/news.gif" width="15" height="13"><a href="http://news.269.net" class="5">����</a> 
      <img border="0" src="http://cx.269.net/bbs/images/soft.gif" width="15" height="13"><a href="http://soft.269.net" class="5">���</a> 
      <img border="0" src="http://cx.269.net/bbs/images/netclub.gif" width="15" height="13"><a href="http://202.100.52.7/netclub/" class="5">��������</a> 
      <img border="0" src="http://cx.269.net/bbs/images/uo.gif" width="15" height="13"><a href="http://love.269.net" class="5">����</a> 
      <img border="0" src="http://cx.269.net/bbs/images/stock.gif" width="15" height="13"><a href="http://stock.269.net" class="5">���з���</a></td>
  </tr>
  <tr> 
    <td bgcolor="#006699" height="1" colspan="2"><img border="0" src="images/1space.gif" width="1" height="1"></td>
  </tr>
</table>
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td valign="bottom" align="right">
      <table width=750 border=0 cellpadding=0 cellspacing=0>
        <tr> 
          <td background="../school_images/nav2_01.gif" width="395">&nbsp;</td>
          <td> <a href="../register.htm" target="_blank"><img src="../school_images/nav2_02.gif" width=75 height=22 border="0"></a></td>
          <td> <img src="../school_images/nav2_03.gif" width=25 height=22></td>
          <td> <a href="../schoollist.asp" target="_blank"><img src="../school_images/nav2_04.gif" width=59 height=22 border="0"></a></td>
          <td> <img src="../school_images/nav2_05.gif" width=30 height=22></td>
          <td> <a href="../help.htm" target="_blank"><img src="../school_images/nav2_06.gif" width=58 height=22 border="0"></a></td>
          <td> <img src="../school_images/nav2_07.gif" width=24 height=22></td>
          <td> <a href="http://202.100.0.104/cgi-bin/chat/index.pl" target="_blank"><img src="../school_images/nav2_08.gif" width=84 height=22 border="0"></a></td>
        </tr>
      </table>
        </td>
  </tr>
  <tr> 
    <td bgcolor="#006699" height="1"></td>
  </tr>
</table>
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="left" valign="top" bgcolor="#C7E4F1" width="165"><br>
      <table width=153 border=0 cellpadding=0 cellspacing=0>
        <tr> 
          <td> <a href="javascript:ClassBBSList();"><img src="../school_images/nav_01.gif" width=153 height=24 border="0"></a></td>
        </tr>
        <tr> 
          <td> <img src="../school_images/nav_02.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td> <a href="javascript:PopWindowUpload();"><img src="../school_images/nav_upload.gif" width="153" height="22" border="0"></a></td>
        </tr>
        <tr> 
          <td> <img src="../school_images/nav_04.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td height="22"><a href="javascript:searchmate();"><img src="../school_images/nav_searchmates.gif" width="153" height="22" border=0></a></td>
        </tr>
        <tr> 
          <td> <img src="../school_images/nav_04.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td height="22"><a href="javascript:PopWindowEmailList();"><img src="../school_images/nav_emaillist.gif" width="153" height="22" border="0"></a></td>
        </tr>
        <tr> 
          <td height=3><img src="../school_images/nav_06.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td><a href="javascript:manage();"><img src="../school_images/nav_manage.gif" width="153" height="22" border=0></a></td>
        </tr>
        <tr> 
          <td><img src="../school_images/nav_02.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td><a href="javascript:ClassBBSList2();"><img src="../school_images/nav_srliuyan.gif" width="153" height="22" border="0"></a></td>
        </tr>
        <tr> 
          <td><img src="../school_images/nav_02.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td><a href="javascript:personalinfo('');"><img src="../school_images/person_info.gif" width="153" height="22" border=0></a></td>
        </tr>
        <tr> 
          <td><img src="../school_images/nav_02.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td><a href="javascript:newclass();"><img src="../school_images/nav_newclass.gif" width="153" height="22"  border=0></a></td>
        </tr>
        <tr> 
          <td><img src="../school_images/nav_02.gif" width=153 height=3 ></td>
        </tr>
        <tr> 
          <td><a href="javascript:modifyinfo();"><img src="../school_images/nav_modifyinfo.gif" width="153" height="22" border=0></a></td>
        </tr>
        <tr> 
          <td><img src="../school_images/nav_02.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td><a href="javascript:PopWindowEmailMates();"><img src="../school_images/nav_zhongzhi.gif" width="153" height="22" border="0"></a></td>
        </tr>
        <tr> 
          <td><img src="../school_images/nav_02.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td><a href="javascript:personallogout();"><img src="../school_images/nav_zhuxiao.gif" width="153" height="22" border=0></a></td>
        </tr>
        <tr> 
          <td><img src="../school_images/nav_02.gif" width=153 height=3></td>
        </tr>
        <tr> 
          <td><a href="../classlogout.asp"><img src="../school_images/nav_09.gif" width=153 height=24 border=0></a></td>
        </tr>
      </table>
    </td>
    <td bgcolor="#C7E4F1" align="center" valign="top" width="22"> <br>
      <table width="22" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="80" background="../school_images/bg1.gif" bgcolor="#ECF5FB" align="center"><a href="class_index.asp" class=3>�༶��Ϣ</a></td>
        </tr>
        <tr> 
          <td height="80" background="../school_images/bg1.gif" align="center"><a href="ClassBBSList.asp" class=3>���Բ�</a></td>
        </tr>
        <tr> 
          <td height="80" background="../school_images/bg1.gif" align="center"><a href="ClassPhotoList.asp" class=3>�������</a></td>
        </tr>
        <tr> 
          <td height="80" background="../school_images/bg1.gif" align="center"><a href="class_phonebook.asp" class=3>ͨѶ¼</a></td>
        </tr>
        <tr> 
          <td height="80" background="../school_images/bg.gif" align="center"><a href="class_t_phonebook.asp" class=3>��ʦͨѶ¼</a></td>
        </tr>
      </table>
      <br>
      <br>
    </td>
    <td bgcolor="#ECF5FB" align="center" valign="top" width="563"> <br>
      <%
            
            if Request("Page")<>"" then 
       		Page = CLng(Request("Page"))
    	    end if 
            If Page < 1 Then 
                Page = 1
            end if
            PageSize=5
            sql="select count(*) as a from userjoinclassinfo where classid ='"&curclassid&"' and userstatus='��ʦ'"  
            rs.open SQL,schooldb,1,3
            count=rs("a")
            PageCount=CInt(rs("a")/PageSize+0.5)
	    rs.Close
       %>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
          <form name="form3" method="post" action="class_phonebook.asp">
        <tr> 
          <td height="28" class="bt" width="100">��ʦͨѶ¼|</td>
          <td height="28" class="topic" width="410" align="right" valign="top">��<%=pagecount%>ҳ 
            | <a href="class_phonebook.asp?page=1">��һҳ</a> 
            | <a href="class_phonebook.asp?page=<%=page-1%>">��һҳ</a> 
            | <a href="class_phonebook.asp?page=<%=page+1%>">��һҳ</a> 
            | <a href="class_phonebook.asp?page=<%=pagecount%>">ĩҳ</a> | �� 
            <input type="text" name="page" size="2">
            ҳ 
            <input type="submit" name="Submit22" value="GO">
          </td>
        </tr>
        <tr> 
          <td bgcolor="#B4C7D4" height="1" colspan="2"></td>
        </tr>
        </form>
      </table>
      <br>
      <%
            if page=1 then
	       SQL="select top 5 * from userjoinclassinfo where classid ='"&curclassid&"' and userstatus='��ʦ' order by userid  "
 	    else
 	       SQL="select top 5 * from userjoinclassinfo where classid ='"&curclassid&"' and userstatus='��ʦ' and userid not in (Select top "&Cstr(PageSize*(page-1))&" userid from userjoinclassinfo where classid ='"&curclassid&"' and userstatus='��ʦ' order by userid  ) order by userid" 
            end if
            
            rs.open SQL,schooldb
            while not rs.eof 
            sql1="select * from userinfo where userid='"&rs("userid")&"'"
            rss.open SQL1,schooldb
            if not rss.eof then
               realname=rss("realname")
            end if
            rss.close   
            sql1="select * from usercommunicationinfo where userid='"&rs("userid")&"'"
            rss.open SQL1,schooldb
            if not rss.eof then
      %>      
      <table width="510" cellspacing="1" cellpadding="1" align="center">
        <tr> 
          <td colspan="2" height="20" class="topic">�û�����<%=rss("userid")%></td>
          <td height="20" class="topic">У��������<%=realname%></td>
        </tr>
        <tr class="digi"> 
          <%if rss("isqqshow")=true then%>
          <td colspan="3" height="20">OICQ/ICQ���룺<%=rss("qq")%></td>
          <%else%>
          <td colspan="3" height="20">OICQ/ICQ���룺******</td>
          <%end if%>
        </tr>
        <tr class="digi"> 
          <%if rss("istelephoneshow")=true then%>
          <td width="200" height="20">�̶��绰��<%=rss("telephone")%></td>
          <%else%>
          <td width="200" height="20">�̶��绰��******</td>
          <%end if%>
          <%if rss("ismobileshow")=true then%>
          <td width="200" height="20">�ƶ��绰��<%=rss("mobile")%></td>
          <%else%>
          <td width="200" height="20">�ƶ��绰��******</td>
          <%end if%>
          <%if rss("isbpshow")=true then%>
          <td width="200" height="20">������<%=rss("bp")%></td>
          <%else%>
          <td width="200" height="20">������******</td>
          <%end if%>
          
        </tr>
        <tr> 
          <%if rss("iscashow")=true then%>
          <td colspan="3" height="20" class="topic">ͨ�ŵ�ַ��<%=rss("communicationaddr")%></td>
          <%else%>
           <td colspan="3" height="20" class="topic">ͨ�ŵ�ַ��**********</td>
          <%end if%>
        </tr>
        <tr> 
           <%if rss("ishashow")=true then%>
          <td colspan="3" height="20" class="topic">��ס��ַ��<%=rss("homeaddr")%></td>
          <%else%>
           <td colspan="3" height="20" class="topic">��ס��ַ��**********</td>
          <%end if%>
          
        </tr>
        <tr> 
           <%if rss("iswsshow")=true then%>
          <td colspan="3" height="20" class="topic">������λ��<%=rss("workshop")%></td>
          <%else%>
           <td colspan="3" height="20" class="topic">������λ��**********</td>
          <%end if%>
          
        </tr>
        <tr> 
          
          <%if rss("isemailshow")=true then%>
          <td colspan="3" height="20" class="digi">�����ʼ���<a href="mailto:<%=rss("email")%>" class=4><%=rss("email")%></a></td>
          <%else%>
           <td colspan="3" height="20" class="digi">�����ʼ���**********</td>
          <%end if%>
          
        </tr>
      </table>
      <hr width="510" size="1" noshade align="center">
     <%end if
       rss.close
        rs.movenext
           wend
           rs.close
         %>  
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0" bgcolor="#C7E4F1">
        <form name="form4" method="post" action="class_phonebook.asp">
        <tr> 
          <td height="28" class="topic" width="410" align="right" valign="top">��<%=pagecount%>ҳ 
            | <a href="class_phonebook.asp?page=1">��һҳ</a> 
            | <a href="class_phonebook.asp?page=<%=page-1%>">��һҳ</a> 
            | <a href="class_phonebook.asp?page=<%=page+1%>">��һҳ</a> 
            | <a href="class_phonebook.asp?page=<%=pagecount%>">ĩҳ</a> | �� 
            <input type="text" name="page" size="2">
            ҳ 
            <input type="submit" name="Submit22" value="GO">
          </td>
        </tr>
        </form>
      </table>
      <br>
      <br>
      <br>
    </td>
  </tr>
</table>
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
          <td background="../school_images/hline.gif" height="2"></td>
        </tr>
      </table>
<br>
<br>
<table width="700" border="0" cellspacing="2" cellpadding="2" align="center">
  <tr align="center"> 
    <td height="15" class="topic"><a href="http://www.269.net/us.htm">��˾���</a> 
      | <a href="http://www.269.net/ads.htm">������</a> | <a href="mailto:lei@pub.xaonline.com">��ϵ��ʽ</a> 
      | <a href="http://www.269.net/join.htm">��Ƹ��Ϣ</a></td>
  </tr>
</table>
<table width="500" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr align="center"> 
    <td height="20"><span style="FONT-SIZE: 9pt"><font face="arial">Copyright2000 
      <a
 href="http://WWW.269.net"><font face="Arial, Helvetica, sans-serif">www.269.net</font></a></font><br>
      ��Ȩ���У���԰���繫˾ &nbsp;&nbsp;������κ�����ͽ���,����<a href="mailto:lei@pub.xaonline.com">������ϵ</a>!</span> 
    </td>
  </tr>
</table>
      </body>
</html>
