<script language="jscript">
<!--
  function check(form1)
  {
    var fig=true
	if(form1.user_name.value=="")
	  {
	    window.alert("�������ԱID��")
		fig=false
	  }
    if(form1.user_pass.value=="")
	  {
	    window.alert("�������Ա���룡")
		fig=false
	  }
	return fig
  }  
  
  function check1(form1)
  {
    var fig=true
	if(form1.user_name.value=="")
	  {
	    window.alert("�������ԱID��")
		fig=false
	  }
    if(form1.findpass.value=="")
	  {
	    window.alert("�������һ�����Ŀ��")
		fig=false
	  }
	return fig
  }  
  function confirmdel()
{
  var fig=false
  fig=confirm("��ȷ��Ҫɾ������Ϣ��")
  return fig
}

function del()
{
  var fig=false
  fig=confirm("��ȷ��Ҫִ�иò�����")
  return fig
}

function addnew()
{
  var fig=false
  fig=confirm("��ȷ��Ҫ��Ӹ�����ͺ���")
  return fig
}
function edit()
{
  var fig=false
  fig=confirm("��ȷ��Ҫ�޸ĸ�����ͺ���")
  return fig
}
  
//-->
</script>
