<!--#include file=globals.asp -->
<%

ClassID=Session("ClassID")
'if ClassID="" then ClassID="GoodLuck"
if ClassID="" then response.redirect "Error.asp?Info=�����ȵ�½��лл��"

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
%><HTML><HEAD>
<meta http-equiv="Content-Language" content="zh-cn">
<TITLE>����У��¼ <%=schoolname%> </TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT language=javascript src="img/alumni.js"></SCRIPT>
<LINK href="img/alumni.css" type=text/css rel=stylesheet>


<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=white leftMargin=0 topMargin=5 >
<CENTER>
<!--#include file=../top.htm-->
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
            <td><a class="cnn" href="javascript:login()">
            <img src="img/cr_a22.gif" border="0" width="16" height="16"></a></td>
            <td class="cnn"><a class="cnn" href="javascript:login()">
            <font color="#ffffff">��¼</font></a></td>
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
    <TD align=middle width=157 bgColor=#FBAF0C><IMG height=1 
      src="img/c.gif" width=1></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=761 border=0>
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
           

                <a class="cla1" href="javascript:manage();">                      <font color="#FF6600">
                �༶����</font></a></TD>
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
    <TD vAlign=top align=middle width=620 
    background=img/cla_bg2.gif>
      <TABLE cellSpacing=0 cellPadding=0 width=607 border=0>
        <TBODY>
        <TR>
          <TD class=cla1 width=380 height=24>&nbsp;&nbsp;У��¼��ҳ &gt;&gt;<a href="../classlist.asp?schoolid=<%=schoolid%>" class=6 target="_blank"><%=schoolname%></a> 
          &gt;&gt;<%=classname%></TD>
          <TD align=right width=227></TD></TR>
        <TR align=middle>
          <TD colSpan=2 width="607"><IMG height=3 src="img/cla_a1.gif" 
            width=457></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=1 cellPadding=0 width=609 bgColor=#d2d2d2 border=0>
        <TBODY>
        <TR>
          <TD bgColor=#ffffff width="607">
            <div align="center">
              <center>
            <TABLE class=cr4 cellSpacing=0 cellPadding=3 width=623 
            border=0 style="border-collapse: collapse" bordercolor="#111111">
              <TBODY>
              <TR align=middle>
                <TD width="49"><a class="down" href="ClassBBSList.asp"><IMG 
                  height=26 src="img/cla_a40.gif" width=36 
                  border=0></a><BR><a class="cla1" href="ClassBBSList.asp">�༶����</a><BR></TD>
                <TD width="8"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0><BR></TD>
                <TD width="49"><a class="down" href="class_phonebook.asp"><IMG 
                  height=26 src="img/cla_a41.gif" width=36 
                  border=0></a><BR><a class="cla1" href="class_phonebook.asp">
                ��Ա��ַ 
                </a> 
                </TD>
                <TD width="6"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD width="51"><a class="down" href="popwindow_search.html"><IMG 
                  height=26 src="img/cla_a42.gif" width=36 
                  border=0></a><BR><a class="cla1" href="popwindow_search.html">
                У�Ѳ�ѯ</a> 
                </TD>
                <TD width="5"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD width="48"><a class="down" href="PopWindowEmailList.asp"><IMG 
                  height=26 src="img/cla_a43.gif" width=36 
                  border=0></a><BR><a class=cla1 href="PopWindowEmailList.asp">
                �ʼ��б�</a></TD>
                <TD width="5"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD width="48">
                <a class="down" href="javascript:manage();"><IMG 
                  height=26 src="img/cla_a44.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" href="javascript:manage();">�༶����</a></TD>
                <TD width="4">
                <IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD width="284">
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                  <tr>
                <TD align="center"><a class="down" href="ClassPhotoList.asp"><IMG 
                  height=26 src="img/cla_a46.gif" width=36 
                  border=0></a><BR><a class="cla1" href="ClassPhotoList.asp">
                �༶��� 
                </a> 
                </TD>
                <TD align="center"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD align="center"><a class="down" href="javascript:PopWindowUpload();"><IMG 
                  height=26 src="img/cla_a45.gif" width=36 
                  border=0></a><BR><a class="cla1" href="javascript:PopWindowUpload();">
                ������Ƭ</a> 
                </TD>
                <TD align="center"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD align="center"><a class="down" href="popwindow_modifyinfo.asp"><IMG 
                  height=26 src="img/cla_a47.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" href="popwindow_modifyinfo.asp">�޸�����</a> 
                </TD>
                <TD align="center"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD align="center">
                <a class="down" href="popwindow_modifyinfo.asp"><IMG 
                  height=26 src="img/cla_a48.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" href="popwindow_modifyinfo.asp">������Ϣ</a></TD>
                <TD align="center"><IMG height=26 src="img/cla_a33.gif" width=5 
                  border=0></TD>
                <TD align="center"><a class="down" href="../find_class1.asp"><IMG height=26 
                  src="img/cla_a49.gif" width=36 
                  border=0></a><BR>
                <a class="cla1" id="t0" target="_blank" href="../find_class1.asp">
                ע���°�</a></TD>
                  </tr>
                </table>
                </TD></TR>
              </TBODY></TABLE></center>
            </div>
          </TD></TR>
        </TBODY></TABLE>
      <TABLE cellSpacing=3 cellPadding=0 width=453 border=0>
        <TBODY>
        <TR>
          <TD align=right width=76>��</TD>
          <TD width="368">
          ��</TD></TR>
        </TBODY></TABLE>
      <TABLE cellSpacing=3 cellPadding=0 width=600 border=0 style="border-collapse: collapse" bordercolor="#111111" height="48">
        <TBODY>
        <TR>
          <TD height=1 width="594" align="center">
          <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CECECE" width="590">
            <tr>
              <td align="middle" bgColor="#f7f7f7" width="588">
              <table height="28" cellSpacing="0" cellPadding="0" border="0">
                <tr>
                  <td width="31">
                  <img src="img/clam_071.gif">
                  </td>
                  <td><strong><%=classname%>�ʼ��б�</strong></td>
                </tr>
              </table>
              </td>
            </tr>
          </table>
          </TD></TR>
        <TR>
          <TD height=12 width="594">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" align="center"> <br>
      <table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" class="topic" width="273"><font color="#CC0000">��Ա</font></td>
          <td height="20" class="digi" width="317" align="left" valign="middle"><font color="#CC0000">EAML��ַ</font></td>
        </tr>
        <tr> 
          <td bgcolor="#B4C7D4" height="1" colspan="2" width="590"></td>
        </tr>
      </table>
      <table width="590" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
<%
set rs = server.createobject("ADODB.recordset")
sql="select U2.RealName,U3.Email from UserJoinClassInfo as U1,UserInfo as u2,UserCommunicationInfo as u3  where U1.UserID=U2.UserID and U1.UserID=U3.UserID and  ClassID='"&ClassID&"'"
rs.Open SQL,DBParams
while not rs.eof
%>
        <tr> 
          <td width="315" height="20" valign="bottom"><%=rs("RealName")%></td>
          <td width="360" valign="bottom" class="en"><a href="mailto:<%=rs("Email")%>" class=4><%=rs("Email")%></a></td>
        </tr>
<%
   rs.MoveNext
wend
rs.close
set rs=nothing
   
%>       
      </table>
             
              ��</td>
            </tr>
          </table>
          </TD></TR>
        </TBODY></TABLE>
      <TABLE class=cr4 cellSpacing=0 cellPadding=5 width=430 align=center 
      border=0>
        <TBODY>
        <TR>
          
</TD></TR></TBODY></TABLE>
      </TD>
    <TD style="BORDER-LEFT: #fbaf0c 7px solid" vAlign=top align=middle width=4 
    bgColor=#FBAF0C>
      ��</TD></TR></TBODY></TABLE>
<TABLE height=1 cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR>
    <TD vAlign=top bgColor=#fbaf0c height="1"><IMG height=1 src="img/c.gif" 
      width=1></TD>
    <TD align=middle width=157 bgColor=#FBAF0C height="1">
    ��</TD></TR></TBODY></TABLE>
</CENTER>
<!--#include file=../end.htm--></BODY></HTML>