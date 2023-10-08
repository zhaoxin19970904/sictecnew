<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Write Back to Client</title>
</head>

<body>
<%
		id=request.QueryString("id")
		sql="select * from inquiry where id="&id&""
		set rs=conn.execute(sql)
		set rs1=conn.execute("select name from products where id="&rs("productid")&"")
		productname=rs1(0)
		rs1.close
		set rs1=nothing
%>
<form action="savebackinquiry.asp?id=<%=id%>" method="post" name="form1">

  <table width="435" height="276" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
    <td width="435" height="131" valign="top"> <table width="433" height="179" border="0" cellpadding="0" cellspacing="0">
          <tr> 
          <td colspan="2"> <div align="center"><strong><font color="#003333">Ask</font></strong></div></td>
        </tr>
        <tr> 
          <td width="68" height="23" valign="top">Subject:</td>
          <td width="506" valign="top"><input name="ask_subject" type="text" id="ask_subject" value="<%=rs("subject")%>" size="40" readonly="">
              <input type="hidden" name="member" value="<%=rs("memberid")%>"></td>
        </tr>
        <tr> 
          <td height="22" valign="top">Date:</td>
          <td valign="top"><input name="ask_date" type="text" id="ask_date" value="<%=rs("date")%>" size="40" readonly=""></td>
        </tr>
        <tr> 
          <td valign="top">Product:</td>
          <td valign="top"><input name="ask_product" type="text" id="ask_product" value="<%=productname%>" size="40" readonly=""></td>
        </tr>
        <tr> 
          <td height="96" valign="top">To Know:</td>
          <td valign="top">
		    <%
	toknow=rs("toknow")
	apart=split(toknow,",")
	i=0
	str=""
	do while i<=ubound(apart)
	 if apart(i)="" then
	   elseif apart(i)<>"" then
	    str=str&apart(i)&","
	  end if
	  i=i+1
	loop
	%>
<textarea name="ask_toknow" cols="40" rows="5" readonly id="ask_toknow">
<%=trim(mid(str,1,len(str)-1))%>
		  </textarea></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="1" bgcolor="#224797"></td>
  </tr>
  <tr> 
    <td valign="top"> <table width="434" height="178" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="22" colspan="2"> <div align="center"><font color="#003366"><strong>Answer</strong></font></div></td>
          </tr>
          <tr> 
            <td width="67" height="24" valign="top">Subject:</td>
            <td width="367" valign="top"><input name="back_subject" type="text" id="back_subject" size="40"></td>
          </tr>
          <tr> 
            <td height="25" valign="top">Date:</td>
            <td valign="top"><input name="textfield5" type="text" size="40" value="<%=date()%>" readonly=""></td>
          </tr>
          <tr> 
            <td height="23" valign="top">Content:</td>
            <td valign="top"><textarea name="back_content" cols="40" rows="5" id="back_content" ></textarea></td>
          </tr>
          <tr> 
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2"> <div align="center"> 
                <input type="submit" name="Submit2" value="提交">
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                <input type="reset" name="Submit" value="重置">
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; [ 
				<a href="#" onclick="javascript:history.back()">返 回</a> ]</div></td>
          </tr>
        </table></td>
  </tr>
</table>
</form>
</body>
</html>
