
<%@ Language=VBScript %>
<%Response.buffer=true%>
<!--#include file=globals.asp -->
<script language="javascript">
function selectclass(curid){

            if (curid!=0) {
                window.location.href="class_index.asp?selectclassid="+curid
        }
    }
function personalinfo(userid){


        window.open("popwindow_perinfo.asp?userID="+userid,"top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=580,height=400")

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
            function PopWindowUpload(){


        window.open("PopWindowUpload.asp","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=570,height=380")

    }

</script>
<%

userid=session("userid")
 realname=session("realname")
set rs = createobject("ADODB.recordset")
selectclassid=trim(request("selectclassid"))
act=trim(request("act"))
if act="del" then
   sql="select * from classinfo where classid='"&session("classid")&"'"
   rs.open SQL,schooldb
   if not rs.eof then
      monitor=rs("monitor")
      monitor1=rs("monitor1")
   end if
   rs.close
   if userid=monitor then
      response.redirect "../error.asp?info=�Բ������Ǳ�������೤��������ע���Լ�����ݣ�"
   end if
   if userid= monitor1 then
      sql="select * from classinfo where classid='"&session("classid")&"'"
      rs.open SQL,schooldb,1,3
      if not rs.eof then
         rs("monitor1")=""
         rs.update
      end if
      rs.close
   end if
   sql="delete from userjoinclassinfo where userid='"&userid&"' and classid='"&session("classid")&"'"
   rs.open SQL,schooldb,1,3
   SQL = "select * from userjoinclassinfo where userid='"&userid&"' order by joindate"
   rs.open SQL,schooldb
   if not rs.eof then
      response.redirect "class_index.asp?selectclassid="&rs("classid")
   else
      Session.Abandon
      response.redirect "../success.asp?info=���Ѿ�ע�����ڸð༶����ݣ�����������û�м����κ�һ���༶��<a href='http://www.sun2003.com/school'>�������</a>,��½������ע��༶��"
   end if
   rs.close
end if
if selectclassid<>"" then
   sql="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&selectclassid&"'"
   rs.open SQL,schooldb
   if rs.eof then
      response.redirect "../error.asp?info=�Բ��������Ǳ����Ա��Ȩ���룡"
   end if
   rs.close
   session("classid")=selectclassid
   sql="select * from online where userid='"&userid&"'"
   rs.open SQL,schooldb,1,3
   if not rs.eof then
      rs("classid")=selectclassid

      rs.update
   end if
   rs.close
   sql="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&selectclassid&"'"
   rs.open SQL,schooldb,1,3
   if not rs.eof then
     if session("loginflag")="" then
        rs("logintimes")=rs("logintimes")+1
        session("loginflag")=true
     end if
     rs("lastlogintime")=now
     rs.update
   end if
   rs.close
end if


curclassid=Session("classid")
if userid="" then
  response.redirect "../error.asp?info=�Բ������Ѿ������ˣ������½��룡"
end if

set rss = createobject("ADODB.recordset")
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
  classname=rs("classname")
  monitor=rs("monitor")
  monitor1=rs("monitor1")
  tonggao=rs("tonggao")
   regdate=rs("regdate")
   enterdate=rs("enterdate")
end if
rs.close

schoolid=left(curclassid,8)
sc=schoolid
sql="select * from schoolinfo where schoolid='"&schoolid&"'"
rs.open SQL,schooldb
if not rs.eof then
   schoolname=rs("schoolname")
end if
rs.close
sql="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
   logintimes=rs("logintimes")
   userstatus=rs("userstatus")

end if
rs.close
sql="select count(*) as a from userjoinclassinfo where userid='"&userid&"'"
rs.open SQL,schooldb,1,3
joinnum=rs("a")
rs.close
sql="select count(*) as b from userjoinclassinfo where classid='"&curclassid&"' and userstatus='��Ա'"
rs.open SQL,schooldb,1,3
membernum=rs("b")
rs.close
sql="select count(*) as c from userjoinclassinfo where classid='"&curclassid&"' and userstatus='��ʦ'"
rs.open SQL,schooldb,1,3
teachernum=rs("c")
rs.close
sql= "Select count(*) as c from message where ClassID='"&curclassid&"' and Deleted=0 and Hits=0 and ToUserID='"&userid&"'"
rs.open SQL,schooldb,1,3
MsgNum=rs("c")
rs.close
%>
<%
set rs=server.createobject("adodb.recordset")
ClassID=Session("ClassID")
%>
<HTML><HEAD>
<meta http-equiv="Content-Language" content="zh-cn">
<TITLE>����У��¼ <%=schoolname%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT language=javascript src="img/alumni.js"></SCRIPT>
<LINK href="img/alumni.css" type=text/css rel=stylesheet>

<script language="javascript">
function Del(URL)
   {
    if (confirm("��ȷ��Ҫɾ����"))
       {
        //window.open("notebook.asp?curID="+curid+"&act=del","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=640,height=400")
       window.location.href=URL
       }
   }

</script>
<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=white leftMargin=0 topMargin=5 >
<CENTER>
<TABLE cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR>
    <TD vAlign=top bgColor=#fbaf0c height=7 colspan="2" width="760">
    <table cellSpacing="0" cellPadding="0" width="760" border="0" height="1">
      <tr>
        <td width="141" bgcolor="#FFFFFF" height="1" valign="bottom">
        <img src="img/cr_a1.gif" width="133" height="6"></td>
        <td width="72" bgcolor="#FFFFFF" height="1">��</td>
        <td width="98" bgcolor="#FFFFFF" height="1">��</td>
        <td width="105" bgcolor="#FFFFFF" height="1">��</td>
        <td width="108" bgcolor="#FFFFFF" height="1">��</td>
        <td width="102" bgcolor="#FFFFFF" height="1">��</td>
        <td vAlign="top" width="134" bgcolor="#FFFFFF" height="1">��</td>
      </tr>
      <tr>
        <td width="141" height="32" background="img/cr_a31.gif">
        <img src="img/cr_a2.gif" border="0" width="141" height="32"></td>
        <td width="72" background="img/cr_a31.gif" height="32">
        ��</td>
        <td width="98" height="32" background="img/cr_a31.gif">
        ��</td>
        <td width="105" height="32" background="img/cr_a31.gif">
        ��</td>
        <td width="108" height="32" background="img/cr_a31.gif">
        ��</td>
        <td width="102" background="img/cr_a31.gif" height="32">
        ��</td>
        <td vAlign="top" width="134" background="img/cr_a32.gif" height="32">
        <table height="26" cellSpacing="0" cellPadding="0" width="120" border="0">
          <tr>
            <td>
            <img src="img/cr_a22.gif" border="0" width="16" height="16"></td>
            <td class="cnn">
            <font color="#ffffff">��¼</font></td>
            <td width="12">
            <img src="img/cr_a23.gif" width="2" height="16"></td>
            <td><a class="cnn" href="../classlogout.asp">
            <img src="img/cr_a22.gif" border="0" width="16" height="16"></a></td>
            <td class="cnn"><a class="cnn" href="../classlogout.asp">
            <font color="#ffffff">ע��</font></a></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </TD>
    </TR>
  <TR>
    <TD vAlign=top bgColor=#fbaf0c height=7 width="607"><IMG height=1 
      src="img/c.gif" width=1></TD>
    <TD align=middle width=157 bgColor=#099acb><IMG height=1 
      src="img/c.gif" width=1></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR>
    <TD class=bg1 style="BORDER-LEFT: #fbaf0c 7px solid" vAlign=top align=middle 
    width=123>
      <TABLE height=5 cellSpacing=0 cellPadding=0 width=100 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=1 src="img/c.gif" 
        width=1></TD></TR></TBODY></TABLE><LINK href="img/photo.css" 
      type=text/css rel=stylesheet>
      <TABLE height=22 cellSpacing=2 cellPadding=0 width=114 
      background=img/cla_bg3.gif border=0>
        <TBODY>
        <TR>
          <TD vAlign=bottom align=middle>�ҵ�У��¼</TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=114 border=0 align=middle>
        <TBODY align=middle >
   
          &nbsp;&nbsp;����:<font color="#FF6600"><%=userid%></font></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=114   height=362 border=0>
        <TBODY>
        <TR vAlign=top>
          <TD height="14" width="114" colspan="3" align="center">
           

                      </TD></TR>
        <TR vAlign=top>
          <TD height="7" width="114" colspan="3">
           

                      </TD></TR>
        <tr>
          <TD height="18" align="center" width="103" colspan="2">
           

                      ����������</TD>
          <TD height="18" align="center" width="11">
           

                      </TD>
        </tr>
        <tr>
          <TD height="22" align="center" valign="top" width="114" colspan="3">
           

                      <a class="cla1" href="popwindow_modifyinfo.asp" style="text-decoration: none">
                      <font color="#FF6600">������Ϣ</font></a></TD>
        </tr>
        <tr>
          <TD height="18" align="center" width="103" colspan="2">
           

                      ���༶��Ϣ</TD>
          <TD height="18" align="center" valign="top" width="11">
           

                      </TD>
        </tr>
        <tr>
          <TD height="112" align="center" valign="top" width="114" colspan="3">
           

                      <a class="cla1" href="ClassBBSList.asp">
                      <font color="#FF6600">�༶����</font></a><br>
                      <a class="cla1" href="class_phonebook.asp">
                      <font color="#FF6600">��Ա��ַ</font></a><br>
                      <a class=cla1 href="PopWindowEmailList.asp">
                      <font color="#FF6600">�ʼ��б�</font></a><br>
                      <a class="cla1" href="ClassPhotoList.asp">
                      <font color="#FF6600">�༶���</font></a><br>
                      <a class="cla1" href="javascript:PopWindowUpload();">
                      <font color="#FF6600">������Ƭ</font></a><br>
                      <a class="cla1" href="tj.asp" style="text-decoration: none">
                      <font color="#FF6600">����ͳ��</font></a></TD>
        </tr>
        <tr>
          <TD height="18" align="center" width="91">
           

                      ������Ա</TD>
          <TD height="18" align="center" width="12">
           

                      </TD>
          <TD height="18" align="center" width="11">
           

                      </TD>
        </tr>
        <tr>
          <TD height="22" align="center" valign="top" width="114" colspan="3">
           

                <a class="cla1" href="javascript:manage();">                      <font color="#FF6600">�༶����</font></a></TD>
        </tr>
        <tr>
          <TD height="17" align="center" width="103" colspan="2">
           

                      �������Ϣ</TD>
          <TD height="17" align="center" valign="top" width="11">
           

                      </TD>
        </tr>
        <tr>
          <TD height="114" align="center" valign="top" width="114" colspan="3">
           

                      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="103">
                        <tr>
                          <td width="29%" height="103">��</td>
                          <td width="71%" height="103" valign="top">
                          <a class="cla1" href="../register.htm" style="text-decoration: none">
                          <font color="#FF6600">���û�ע��</font></a><br>
                          <a class="cla1" href="../schoollist.asp" style="text-decoration: none">
                          <font color="#FF6600">���ѧУ</font></a><br>
                          <a class="cla1" href="../find_class1.asp" style="text-decoration: none">
                          <font color="#FF6600">ע���°༶</font></a></td>
                        </tr>
                      </table>
          </TD>
        </tr>
        </TBODY></TABLE></TD>
    <TD vAlign=top align=middle width=466 
    background=img/cla_bg2.gif>
      <TABLE cellSpacing=0 cellPadding=0 width=457 border=0>
        <TBODY>
        <TR>
          <TD class=cla1 width=380 height=24>&nbsp;&nbsp;У��¼��ҳ &gt;&gt;<a href="../classlist.asp?schoolid=<%=schoolid%>" class=6 target="_blank"><%=schoolname%></a> 
          &gt;&gt;<a href="class_index.asp"><%=classname%></a></TD>
          <TD align=right width=77></TD></TR>
        <TR align=middle>
          <TD colSpan=2><IMG height=3 src="img/cla_a1.gif" 
            width=457></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=1 cellPadding=0 width=447 bgColor=#d2d2d2 border=0>
        <TBODY>
        <TR>
          <TD bgColor=#ffffff colSpan=2 width="445">
            <TABLE height=26 cellSpacing=0 cellPadding=4 width=100 border=0>
              <TBODY>
              <TR>
                <TD><FONT color=#0066cc><STRONG>&nbsp;<FONT 
                  color=#246f00>�༶����</FONT></STRONG></FONT></TD></TR></TBODY></TABLE>
            <TABLE class=cr4 cellSpacing=0 cellPadding=3 width=400 align=center 
            border=0>
              <TBODY>
              <TR align=middle>
                <TD><a class="down" href="ClassBBSList.asp"><IMG 
                  height=26 src="img/cla_a40.gif" width=36 
                  border=0></a><BR><a class="cla1" href="ClassBBSList.asp">�༶����</a><BR></TD>
                <TD><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0><BR></TD>
                <TD><a class="down" href="class_phonebook.asp"><IMG 
                  height=26 src="img/cla_a41.gif" width=36 
                  border=0></a><BR><a class="cla1" href="class_phonebook.asp">��Ա��ַ 
                </a> 
                </TD>
                <TD><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD><a class="down" href="popwindow_search.html"><IMG 
                  height=26 src="img/cla_a42.gif" width=36 
                  border=0></a><BR><a class="cla1" href="popwindow_search.html">У�Ѳ�ѯ</a> 
                </TD>
                <TD><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD><a class="down" href="PopWindowEmailList.asp"><IMG 
                  height=26 src="img/cla_a43.gif" width=36 
                  border=0></a><BR><a class=cla1 href="PopWindowEmailList.asp">�ʼ��б�</a></TD>
                <TD><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD>
                <a class="down" href="javascript:manage();"><IMG 
                  height=26 src="img/cla_a44.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" href="javascript:manage();">�༶����</a></TD></TR>
              <TR align=middle>
                <TD><a class="down" href="ClassPhotoList.asp"><IMG 
                  height=26 src="img/cla_a46.gif" width=36 
                  border=0></a><BR><a class="cla1" href="ClassPhotoList.asp">�༶��� 
                </a> 
                </TD>
                <TD><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD><a class="down" href="javascript:PopWindowUpload();"><IMG 
                  height=26 src="img/cla_a45.gif" width=36 
                  border=0></a><BR><a class="cla1" href="javascript:PopWindowUpload();">������Ƭ</a> 
                </TD>
                <TD><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD><a class="down" href="popwindow_modifyinfo.asp"><IMG 
                  height=26 src="img/cla_a47.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" href="popwindow_modifyinfo.asp">�޸�����</a> 
                </TD>
                <TD><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD><a class="down" href="popwindow_modifyinfo.asp"><IMG 
                  height=26 src="img/cla_a48.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" href="popwindow_modifyinfo.asp">������Ϣ</a></TD>
                <TD><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD><a class="down" href="../find_class1.asp"><IMG height=26 
                  src="img/cla_a49.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" id="t0" target="_blank" href="../find_class1.asp">ע���°�</a></TD></TR></TBODY></TABLE></TD></TR>
        <TR>
          <TD width=230 bgColor=#f3f3f3>
            <TABLE height=26 cellSpacing=0 cellPadding=4 width=100 border=0>
              <TBODY>
              <TR>
                <TD><FONT color=#0066cc><STRONG><FONT 
                  color=#246f00>&nbsp;�༶��Ϣ</FONT></STRONG></FONT></TD></TR></TBODY></TABLE>
            <TABLE class=cla1 cellSpacing=0 cellPadding=0 width=229 border=0>
              <TBODY>
              <TR>
                <TD align=right width=62><FONT color=#46a718>�ࡡ����:</FONT></TD>
                <TD class=cr2 width=167 colspan="3"><%=classname%></TD></TR>
              <TR>
                <TD align=right width=62><FONT color=#46a718>��ѧ���:</FONT></TD>
                <TD width=80><FONT color=#4b4b4b><%=enterdate%></FONT></TD>
                <TD width=55><FONT color=#46a718>�� Ա ��:</FONT></TD>
                <TD width=32><%=membernum%>��</TD>
                </TR>
              <TR>
                <TD align=right width="62"><font color="#46A718">����</font><FONT color=#46a718>ʱ��:</FONT></TD>
                <TD align=left width="167" colspan="3"><%=regdate%> </TD></TR>
              <TR>
                <TD align=right width="62"><FONT color=#46a718>�� ʼ ��:</FONT></TD>
                <TD align=left width="167" colspan="3"><A class=cla2 href="javascript:personalinfo('<%=trim(monitor)%>');"><%=monitor%></a> 
                  </A></TD></TR>
              </TBODY></TABLE></TD>
          <TD width=214 bgColor=#E2F9FF>
            <TABLE height=26 cellSpacing=0 cellPadding=4 width=100 border=0>
                            <TBODY>
              <TR>
                <TD><FONT color=#0066cc><STRONG><FONT 
                  color=#246f00>&nbsp;����У��ע��</FONT></STRONG></FONT></TD></TR></TBODY></TABLE>
            &nbsp;����ͬѧ��֪������༶��<BR>&nbsp;��췢�������ʼ����ټ����Ǽ����! 
            <BR>&nbsp;�뷢�������! <BR> 
            <INPUT type=radio value=2 name=sendtype checked>�ʼ� &nbsp; &nbsp;<a href="../error.asp?info=�Բ��𣬸÷���Ϊ�շѷ��� ��"><img border="0" src="img/cla_a53.gif"></a> </TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=3 cellPadding=0 width=453 border=0>
        <TBODY>
        <TR>
          <TD align=right width=76><FONT color=#46a718>��վ��ַ:</FONT></TD>
          <TD width="368">
          <a target="_blank" href="http://www.auce.org" style="text-decoration: none">
          <font color="#FF6600">http://www.auce.org</font></a></TD></TR>
        <TR>
          <TD align=right width="76"><font color="#46A718">��</font><FONT color=#46a718>�³�Ա:</FONT></TD>
          <TD width="368"><%
              sql="select top 3 * from userjoinclassinfo where classid='"&curclassid&"' and userstatus='��Ա' order by joindate desc"
              rs.open SQL,schooldb
              while not rs.eof
            %>


            <a class=cla2 href="javascript:personalinfo('<%=trim(rs("userid"))%>');"><%=rs("userid")%></a>
           <%
            rs.movenext
            wend
            rs.close
            %></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=3 cellPadding=0 width=453 border=0>
        <TBODY>
        <TR>
          <TD bgColor=#D0F3F9 height=24>��<FONT color=#246f00><STRONG>�༶���� 
            </STRONG></FONT>
            <MARQUEE onmouseover=this.stop() onmouseout=this.start() 
            scrollAmount=3 scrollDelay=120 width=350>��<A class=cla2&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; href="http://www.88com.cn"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; target=_blank>Wosou����ϵͳȫ�����������������ӿ�ݺͷ���Ҳ�������Ի�!!</A>&nbsp;&nbsp;&nbsp;</MARQUEE></TD></TR></TBODY></TABLE>
      <TABLE class=cr4 cellSpacing=0 cellPadding=5 width=430 align=center 
      border=0>
        <TBODY>
        <TR>
          <TD class=cla1><%if tonggao<>"" or not isnull(tonggao) then%>
          <td class="topic" valign="top"><%=tonggao%></td>
          <%else%>
          <td class="topic" valign="top"><font color="#4B4B4B">�ˣ���Һã���ӭ����Wosouͬѧ¼��<br>
            ϣ��ͨ������ܷ�������ϵͬ�����Ѻ�ͬ��ͬѧ���������ʹ������ʲô������ĵط������������Ķ�</font><a href="../help.htm" target="_blank"><font color="#FF6600">�����ļ�</font></a><font color="#4B4B4B">�����߸��������ţ�</font><br>
            <font color="#FF6600">���༶���ԡ������ɹ���Ա���߸�����Ա�ڡ��༶�����з�����</font></td>
          <%end if%>
</TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=3 cellPadding=0 width=453 border=0>
        <TBODY>
        <TR>
          <TD bgColor=#D0F3F9 height=24>��<STRONG><FONT color=#246f00>�༶���� 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></STRONG>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <a href="ClassBBSList.asp"><IMG 
            height=18 alt=�༶��Ա�������� src="img/cla_a61.gif" width=59 
            align=absMiddle border=0></a>&nbsp; <a href="ClassBBSList.asp">
          <IMG 
            height=18 alt=�鿴�������� src="img/cla_a65.gif" width=59 
            align=middle border=0></a></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=447 border=0>
        <TBODY>
        <TR bgColor=#f3f3f3>
          <TD align=middle width="100%" height=23>
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="71%" bgcolor="#FFFFFF" height="1">
            <tr>
              <td width="100%" height="1"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="105%">
                <tr>
                  </td>
                </tr>
              </table>
              <div align="center">
                <center>
                <p><td width="100%">��<%

       PageSize=15
    rs.open "Select count(*) as c from message where ClassID='"&ClassID&"' and Deleted=0 "&SqlAdd,DBParams,1,3
    PageCount=CInt(rs("c")/PageSize+0.5)
    rs.Close
    Page = CLng(Request("Page"))
    if Page=""  or isnull(Page) then  page=1
    If Page < 1 Then    Page = 1
    If Page > PageCount Then  Page = PageCount
%>
   
<%


        if page=1 or page=0 then
        SQL="select top "&PageSize&" * from message where ClassID='"&ClassID&"' and Deleted=0 "&SqlAdd&" order by MsgID DESC"
    else
        SQL="select  top "&PageSize&" * from message where ClassID='"&ClassID&"' and Deleted=0 "&SqlAdd&" and "
        SQL=SQL & " MsgID not in (Select top "&Cstr(PageSize*(page-1))&" MsgID from message where "
        SQl=SQL & " ClassID='"&ClassID&"' and Deleted=0 "&SqlAdd&" order by MsgID DESC)"
        SQL=SQL & " order by MsgID DESC"
    end if
    'Response.Write sql&page
    rs.Open SQL,DBParams
   while not rs.eof
%>
      <table width="444" border="0" cellspacing="0" cellpadding="0" height="16">
        <tr>
          <td height="1" width="17" bgcolor="#F3F3F3"><img src="../school_images/topic_inco1.gif" width="15" height="15"></td>
          <td height="25" width="394" class="topic2" bgcolor="#D0F3F9"><font color="#246F00"><%=rs("RegTime")%></font>&nbsp;<%=rs("UserID")%></td>
        </tr>
        
        <tr>
          <td height="25" class="topic"  colspan="2" width="394" height="1"><%=rs("Comment")%></td>
          </tr>
    
</td>
      </table>  <%
      rs.MoveNext
   wend
   rs.close
%></p>
                </center>
              </div>
</td>
            </tr>
          </table>
          </TD>
          </TR>
        </TBODY></TABLE>
      <TABLE cellSpacing=3 cellPadding=0 width=453 border=0>
        <TBODY>
        <TR>
          <TD bgColor=#D0F3F9 height=24>��<STRONG><FONT color=#246f00>�༶���� 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </FONT></STRONG>&nbsp;<a href="ClassBBSList.asp"><IMG 
            height=18 alt=�༶��Ա�������� src="img/cla_a61.gif" width=59 
            align=middle border=0></a>&nbsp; <a href="ClassBBSList.asp">
          <IMG 
            height=18 alt=�鿴�������� src="img/cla_a65.gif" width=59 
            align=middle border=0></a></TD></TR></TBODY></TABLE></TD>
    <TD style="BORDER-LEFT: #fbaf0c 7px solid" vAlign=top align=middle width=157 
    bgColor=#d0f3f9>
      <TABLE cellSpacing=0 cellPadding=0 width=157 border=0 height="41">
        <TBODY>
        <TR>
          <TD align=middle bgColor=#b7f0ff height=24><strong>
          <font color="#0066CC">��ע��������༶</font></strong></TD></TR>
        <TR>
          <TD class=cr4 
          style="PADDING-LEFT: 7px; PADDING-BOTTOM: 4px; PADDING-TOP: 12px" 
          bgColor=#e2f9ff align="center" height="17">
            <select name="select" onChange="javascript:selectclass(this.options[selectedIndex].value);">
              <option selected value=0>ѡ����İ༶</option>
              <%
               sql="select * from userjoinclassinfo where userid='"&userid&"' order by joindate"
               rs.open SQL,schooldb
               while not rs.eof
                 sql1="select * from classinfo where classid='"&rs("classid")&"'"
                 rss.open SQL1,schooldb
                 if not rss.eof then
                  allclassname=rss("classname")
                 end if
                 rss.close
                 allschoolid=left(rs("classid"),8)
                 sql1="select * from schoolinfo where schoolid='"&allschoolid&"'"
                 rss.open SQL1,schooldb
                 if not rss.eof then
                 allschoolname=rss("schoolname")
                 end if
                 rss.close
             %>
             <option value="<%=rs("classid")%>"><%=allschoolname%>--<%=allclassname%></option>
              <%
            rs.movenext
            wend
            rs.close
            %>
            </select><BR></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=157 border=0>
        <TBODY>
        <TR>
          <TD align=middle bgColor=#b7f0ff height=24><strong>
          <font color="#0066CC">����ͬѧ(�Ƽ�)</font></strong></TD></TR>
        <TR>
          <TD class=cr4 
          style="PADDING-LEFT: 7px; PADDING-BOTTOM: 4px; PADDING-TOP: 12px" 
          bgColor=#e2f9ff width="100%" height="100" align="center"><form name="form1" method="post" action="popwindow_searchout.asp" target="_blank">ͬѧ����:<input type="text" name="realname" size="11"><br>
          ����ѧУ:<input type="text" name="schoolname" size="11"><br>
          ���ڰ༶:<input type="text" name="classname" size="11"><br>
          ��ѧ���:<select name=enterdate>
                <option value="0" selected>��ѡ��</option>
                <option value=1970>1970��</option>
                <option value=1971>1971��</option>
                <option value=1972>1972��</option>
                <option value=1973>1973��</option>
                <option value=1974>1974��</option>
                <option value=1975>1975��</option>
                <option value=1976>1976��</option>
                <option value=1977>1977��</option>
                <option value=1978>1978��</option>
                <option value=1979>1979��</option>
                <option value=1980>1980��</option>
                <option value=1981>1981��</option>
                <option value=1982>1982��</option>
                <option value=1983>1983��</option>
                <option value=1984>1984��</option>
                <option value=1985>1985��</option>
                <option value=1986>1986��</option>
                <option value=1987>1987��</option>
                <option value=1988>1988��</option>
                <option value=1989>1989��</option>
                <option value=1990>1990��</option>
                <option value=1991>1991��</option>
                <option value=1992>1992��</option>
                <option value=1993>1993��</option>
                <option value=1994>1994��</option>
                <option value=1995>1995��</option>
                <option value=1996>1996��</option>
                <option value=1997>1997��</option>
                <option value=1998>1998��</option>
                <option value=1999>1999��</option>
                <option value=2000>2000��</option>
                <option value=2001>2001��</option>
                <option value=2002>2002��</option>
                <option value=2003>2003��</option>
              </select><br>
          <br>
              <input type="submit" name="Submit" value="��ʼ����"></form></TD></TR></TBODY></TABLE>��</TD></TR></TBODY></TABLE>
<TABLE height=7 cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR>
    <TD vAlign=top bgColor=#fbaf0c><IMG height=1 src="img/c.gif" 
      width=1></TD>
    <TD align=middle width=157 bgColor=#d0f3f9>
    <a href="javascript:personallogout();">
    <IMG 
      src="img/cla_quit.gif" border=0 width="90" height="18"></a></TD></TR></TBODY></TABLE>
</CENTER>
<!--#include file=../end.htm--></BODY></HTML>