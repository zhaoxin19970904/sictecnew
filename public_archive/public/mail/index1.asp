<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Ԥ���Ƶ�</title>
<style type="text/css">
.test{
  font-size:12px;
}
</style>
</head>
<body>
<form name="form1" method="post" action="script/mail1.asp" onsubmit="return check(this)">
  <table width="607" height="232" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td bgcolor="#CCCCCC"><font color="#000099"> &nbsp;&nbsp;�� �� �� Ϣ</font></td>
    </tr>
    <tr> 
      <td height="120"> <table width="500" height="113" border="0" align="center" cellpadding="0" cellspacing="0" class="test">
          <tr> 
            <td width="79" height="16">����������</td>
            <td width="421"><input name="contact_name" type="text" id="contact_name"> <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">���ĵ�λ��</td>
            <td><input name="company" type="text" id="company"> <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">����ְλ��</td>
            <td><input name="place" type="text" id="place"> </td>
          </tr>
          <tr> 
            <td height="16">��ϵ�绰��</td>
            <td><input name="contact_tel" type="text" id="contact_tel"> <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">E-mail��</td>
            <td><input name="e_mail" type="text" id="e_mail"> <font color="#FF0000">*</font></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC"><font color="#000099"> &nbsp;&nbsp;�� �� �� ��</font></td>
    </tr>
    <tr> 
      <td><table width="500" height="173" border="0" align="center" cellpadding="0" cellspacing="0" class="test">
          <tr> 
            <td width="79" height="16">�Ƶ�Ҫ��</td>
            <td width="421"><input name="shop_ask_city1" type="text" id="shop_ask_city1" size="50"> </td>
          </tr>
          <tr> 
            <td height="16">����һ��</td>
            <td><input name="city1" type="text" id="city1" size="15"> <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ꼶��</td>
            <td><input name="shop1_level" type="radio" value="����" checked>
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop1_level" value="����">
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop1_level" value="����">
              ���� <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ������</td>
            <td> <table width="166" height="77" border="0" cellpadding="0" cellspacing="0" class="test">
                <tr> 
                  <td width="53">���˷���</td>
                  <td width="76"><input name="city1_dr" type="text" id="city1_dr" size="8"></td>
                  <td width="18">��</td>
                  <td width="19" rowspan="3"><div align="center"><font color="#FF0000">*</font></div></td>
                </tr>
                <tr> 
                  <td>˫�˷���</td>
                  <td><input name="city1_sr" type="text" id="city1_sr" size="8"></td>
                  <td>��</td>
                </tr>
                <tr> 
                  <td>�� &nbsp;����</td>
                  <td><input name="city1_tf" type="text" id="city1_tf" size="8"></td>
                  <td>��</td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC" height="1"></td>
    </tr>
    <tr> 
      <td><table width="500" height="173" border="0" align="center" cellpadding="0" cellspacing="0" class="test">
          <tr> 
            <td width="79" height="16">�Ƶ�Ҫ��</td>
            <td width="421"><input name="shop_ask_city2" type="text" id="shop_ask_city2" size="50"> </td>
          </tr>
          <tr> 
            <td height="16">���ж���</td>
            <td><input name="city2" type="text" id="city2" size="15"> <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ꼶��</td>
            <td><input name="shop2_level" type="radio" value="����" checked>
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop2_level" value="����">
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop2_level" value="����">
              ���� <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ������</td>
            <td> <table width="166" height="77" border="0" cellpadding="0" cellspacing="0" class="test">
                <tr> 
                  <td width="53">���˷���</td>
                  <td width="76"><input name="city2_dr" type="text" id="city2_dr" size="8"></td>
                  <td width="18">��</td>
                  <td width="19" rowspan="3"><div align="center"><font color="#FF0000">*</font></div></td>
                </tr>
                <tr> 
                  <td>˫�˷���</td>
                  <td><input name="city2_sr" type="text" id="city2_sr" size="8"></td>
                  <td>��</td>
                </tr>
                <tr> 
                  <td>�� &nbsp;����</td>
                  <td><input name="city2_tf" type="text" id="city2_tf" size="8"></td>
                  <td>��</td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC" height="1"></td>
    </tr>
    <tr> 
      <td><table width="500" height="173" border="0" align="center" cellpadding="0" cellspacing="0" class="test">
          <tr> 
            <td width="79" height="16">�Ƶ�Ҫ��</td>
            <td width="421"><input name="shop_ask_city3" type="text" id="shop_ask_city3" size="50"> </td>
          </tr>
          <tr> 
            <td height="16">��������</td>
            <td><input name="city3" type="text" id="city3" size="15"> <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ꼶��</td>
            <td><input name="shop3_level" type="radio" value="����" checked>
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop3_level" value="����">
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop3_level" value="����">
              ���� <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ������</td>
            <td> <table width="166" height="77" border="0" cellpadding="0" cellspacing="0" class="test">
                <tr> 
                  <td width="53">���˷���</td>
                  <td width="76"><input name="city3_dr" type="text" id="city3_dr" size="8"></td>
                  <td width="18">��</td>
                  <td width="19" rowspan="3"><div align="center"><font color="#FF0000">*</font></div></td>
                </tr>
                <tr> 
                  <td>˫�˷���</td>
                  <td><input name="city3_sr" type="text" id="city3_sr" size="8"></td>
                  <td>��</td>
                </tr>
                <tr> 
                  <td>�� &nbsp;����</td>
                  <td><input name="city3_tf" type="text" id="city3_tf" size="8"></td>
                  <td>��</td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC" height="1"></td>
    </tr>
    <tr> 
      <td><table width="500" height="173" border="0" align="center" cellpadding="0" cellspacing="0" class="test">
          <tr> 
            <td width="79" height="16">�Ƶ�Ҫ��</td>
            <td width="421"><input name="shop_ask_city4" type="text" id="shop_ask_city4" size="50"> </td>
          </tr>
          <tr> 
            <td height="16">�����ģ�</td>
            <td><input name="city4" type="text" id="city4" size="15"> <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ꼶��</td>
            <td><input name="shop4_level" type="radio" value="����" checked>
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop4_level" value="����">
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop4_level" value="����">
              ���� <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ������</td>
            <td> <table width="166" height="77" border="0" cellpadding="0" cellspacing="0" class="test">
                <tr> 
                  <td width="53">���˷���</td>
                  <td width="76"><input name="city4_dr" type="text" id="city4_dr" size="8"></td>
                  <td width="18">��</td>
                  <td width="19" rowspan="3"><div align="center"><font color="#FF0000">*</font></div></td>
                </tr>
                <tr> 
                  <td>˫�˷���</td>
                  <td><input name="city4_sr" type="text" id="city4_sr" size="8"></td>
                  <td>��</td>
                </tr>
                <tr> 
                  <td>�� &nbsp;����</td>
                  <td><input name="city4_tf" type="text" id="city4_tf" size="8"></td>
                  <td>��</td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC" height="1"></td>
    </tr>
    <tr> 
      <td><table width="500" height="173" border="0" align="center" cellpadding="0" cellspacing="0" class="test">
          <tr> 
            <td width="79" height="16">�Ƶ�Ҫ��</td>
            <td width="421"><input name="shop_ask_city5" type="text" id="shop_ask_city5" size="50"> </td>
          </tr>
          <tr> 
            <td height="16">�����壺</td>
            <td><input name="city5" type="text" id="city5" size="15"> <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ꼶��</td>
            <td><input name="shop5_level" type="radio" value="����" checked>
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop5_level" value="����">
              ���� &nbsp;&nbsp;&nbsp; <input type="radio" name="shop5_level" value="����">
              ���� <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="16">�Ƶ������</td>
            <td> <table width="166" height="77" border="0" cellpadding="0" cellspacing="0" class="test">
                <tr> 
                  <td width="53">���˷���</td>
                  <td width="76"><input name="city5_dr" type="text" id="city5_dr" size="8"></td>
                  <td width="18">��</td>
                  <td width="19" rowspan="3"><div align="center"><font color="#FF0000">*</font></div></td>
                </tr>
                <tr> 
                  <td>˫�˷���</td>
                  <td><input name="city5_sr" type="text" id="city5_sr" size="8"></td>
                  <td>��</td>
                </tr>
                <tr> 
                  <td>�� &nbsp;����</td>
                  <td><input name="city5_tf" type="text" id="city5_tf" size="8"></td>
                  <td>��</td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC"><div align="center">
          <input type="submit" name="Submit" value="ȷ��Ԥ��">
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="reset" name="Submit2" value="��д��Ϣ">
        </div></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
</form>

<div align="center"><a href="admin/index.asp">������������</a> 
  <script language="JScript">
<!--
function check(form1)
{
if (form1.contact_name.value=="")
 {
    alert("��������������!")
	form1.contact_name.focus()
    return false
  }
  if (form1.company.value=="")
 {
    alert("���������ĵ�λ!")
	form1.company.focus()
    return false
  }
if (form1.contact_tel.value=="")
 {
    alert("������������ϵ�绰!")
	form1.contact_tel.focus()
    return false
  }
if (form1.e_mail.value=="")
 {
    alert("����������E-mail!")
	form1.e_mail.focus()
    return false
  }
if (form1.city1.value!="")
 {
  if(form1.city1_dr.value=="" && form1.city1_sr.value=="" && form1.city1_tf.value=="")
  {
    alert("��ȷ���ڳ���һ�����Ƶ�ļ���!")
	form1.city1_dr.focus()
    return false
	}
  }
  if (form1.city2.value!="")
 {
  if(form1.city2_dr.value=="" && form1.city2_sr.value=="" && form1.city2_tf.value=="")
  {
    alert("��ȷ���ڳ��ж������Ƶ�ļ���!")
	form1.city2_dr.focus()
    return false
	}
  }
  if (form1.city3.value!="")
 {
  if(form1.city3_dr.value=="" && form1.city3_sr.value=="" && form1.city3_tf.value=="")
  {
    alert("��ȷ���ڳ����������Ƶ�ļ���!")
	form1.city3_dr.focus()
    return false
	}
  }
  if (form1.city4.value!="")
 {
  if(form1.city4_dr.value=="" && form1.city4_sr.value=="" && form1.city4_tf.value=="")
  {
    alert("��ȷ���ڳ����Ĵ����Ƶ�ļ���!")
	form1.city4_dr.focus()
    return false
	}
  }
  if (form1.city5.value!="")
 {
  if(form1.city5_dr.value=="" && form1.city5_sr.value=="" && form1.city5_tf.value=="")
  {
    alert("��ȷ���ڳ���������Ƶ�ļ���!")
	form1.city5_dr.focus()
    return false
	}
  }
}
//-->
</script>
</div>
</body>
</html>
