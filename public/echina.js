$(document).ready(function() {
	$("#know").insertBefore("#answer-form");
	//$(".qainfowtr").children(".qainfowtrs").children("h3").next(".font11px").removeClass("font11px").addClass("font11px EditQuestion");
	$(".qainfowtr>.qainfowtrs>h3").next(".font11px").removeClass("font11px").addClass("font11px EditQuestion");
	$(".page-content>.content>.answer-edit-form>.answer-comment").children(".box").removeClass("box");
	//login
	var url = location.href; 
	if(url.indexOf("/user/login")!=-1)
	{
		$(".content>.lphn").wrapInner("<div class=\"zhucetshi\"></div>");
		$(".content>.lphn>.zhucetshi>.lujingw").insertBefore(".content>.lphn>.zhucetshi");
		$(".content>.lphn>.zhucetshi>.login-content-form").insertAfter(".content>.lphn>.zhucetshi");
		$(".login-content-form>.lphnc>.titletopl").html("Login to eChinacities.com").after("<div class=\"daoru\">Don't have an ID? <a href=\"/ask/user/register\">Register now</a></div>");
		$(".content>.lphn>.lujingw>.breadcrumb").find("a").eq(1).html("Register");
	}
	//////
	//$(".lphnc>.titletopl").after("<div class=\"daoru\">Don't have an ID? <a href=\"/ask/user/register\">Register now</a></div>");
	$(".page-content>.content>.qainfowt").find(".qainfowtl a:last-child").wrap("<div></div>");
	$(".comment").find(".contentsqaright p:first-child").wrap("<h3></h3>");
	$(".page-content>.content>.qainfowt").find(".bestqa").find(".bestqacon >.bestqacont").find("div").eq(1).removeClass("padding-bottom").addClass("padding-bottom1");
	$("#edit-name-wrapper>.description").html("(punctuation is not allowed except for periods, hyphens, and underscores.)");
	$("#edit-pass-pass2-wrapper>label").html("Re-enter Password:<span title=\"This field is required.\" class=\"form-required\">*</span>");
});