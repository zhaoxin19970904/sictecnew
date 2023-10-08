<script language="jscript">
<!--
  function check(form1)
  {
    var fig=true
	if(form1.user_name.value=="")
	  {
	    window.alert("请输入会员ID！")
		fig=false
	  }
    if(form1.user_pass.value=="")
	  {
	    window.alert("请输入会员密码！")
		fig=false
	  }
	return fig
  }  
  
  function check1(form1)
  {
    var fig=true
	if(form1.user_name.value=="")
	  {
	    window.alert("请输入会员ID！")
		fig=false
	  }
    if(form1.findpass.value=="")
	  {
	    window.alert("请输入找回密码的口令！")
		fig=false
	  }
	return fig
  }  
  function confirmdel()
{
  var fig=false
  fig=confirm("您确定要删除该信息吗？")
  return fig
}

function del()
{
  var fig=false
  fig=confirm("您确定要执行该操作吗？")
  return fig
}

function addnew()
{
  var fig=false
  fig=confirm("您确定要添加该外壳型号吗？")
  return fig
}
function edit()
{
  var fig=false
  fig=confirm("您确定要修改该外壳型号吗？")
  return fig
}
  
//-->
</script>
