<HTML>
<HEAD>
<title>上传照片</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<meta name="web_designner" content="校友录">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0"><!--#include file=../top1.htm-->
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="top"> <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="20" align="left" valign="top" class="topic">上传照片</td>
        </tr>
      </table>
      <table width="510" border="0" cellspacing="1" cellpadding="1">
        <form action="popwindowupload1.asp" method=post>
          <tr> 
            <td width="100" align="right" height="26">照片标题：</td>
            <td> 
              <input type="text" name="Title" size="50" value="无标题">
            </td>
          </tr>
          <tr> 
            <td width="100" align="right" height="26">照片类型：</td>
            <td height="26"> 
              <select name="pictype">
                <option selected>真我风采</option>
                <option>集体合影</option>
                <option>其他</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td width="100" align="right" height="26">照片说明：</td>
            <td> 
              <textarea name="Comment" cols="40" rows="5">不错吧
</textarea>
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" height="26"> 
              <input type="submit" name="Submit2" value="下一步">
              <br>
              <br>
            </td>
          </tr>
        </form>
      </table>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>