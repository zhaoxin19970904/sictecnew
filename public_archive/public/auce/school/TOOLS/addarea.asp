<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#006699">
<table width="75%" border="0" align="center">
  <tr align="center"> 
    <td width="750" height="50"><font color="#FFFFFF">填加地区</font></td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="75%" border="0" align="center" height="200">
  <form name="form1" method="post" action="area.asp">
    <tr align="center"> 
      <td height="25%"><font color="#FFFFFF">选择地区所在的省份：</font></td>
      <td height="25%"> 
        <select name="provinceid">
        <%set rs = createobject("ADODB.recordset")
          sql="select * from province order by seed desc"
          rs.open SQL,schooldb
          while not rs.eof
        %>
        <option value="<%=rs("provinceid")%>"><%=rs("provincename")%></option>
        <%rs.movenext
          wend
          rs.close
        %>  
        </select>
        <font color="#FFFFFF"></font></td>
    </tr>
    <tr align="center"> 
      <td height="25%"><font color="#FFFFFF">所要填加的地区的名称：</font></td>
      <td height="25%"> 
        <input type="text" name="areaname">
        <font color="#FFFFFF"></font></td>
    </tr>
    <tr align="center"> 
      <td height="44"> 
        <input type="submit" name="Submit" value="提交">
      </td>
      <td height="44"> 
        <input type="reset" name="Submit2" value="重写">
      </td>
    </tr>
  </form>
</table>
<p>&nbsp;</p>
</body>
</html>
