var today = new Date();
var days = today.getDate();
var hours = today.getHours();
var monthss = today.getMonth();
var minutes = today.getMinutes();
var dayofweeks = today.getDay();
var pagewidth = window.screen.width;
var oddday = days%2
var leftformat;
if (monthss==6 & (days==7 | days==8 | days==9 | days==14 | days==15 | days==16 | days==21 | days==22 | days==23))
{
leftformat="flash";
leftpicsuspent="http://images.sohu.com/cs/button/shida/8080.gif";
leftlinksuspent="http://goto.sohu.com/goto.php3?code=shida-g51-alumnibanjifly";
leftflashsuspent="http://images.sohu.com/cs/button/tclpc/0610/tclpc8080.swf?clickthru=http://goto.sohu.com/goto.php3?code=tclpc-gz142alumniclassfly";
}
else if (monthss==6 & (days==17 | days==18 | days==19 | days==29))
{
leftformat="gif";
leftpicsuspent="http://images.sohu.com/cs/button/benq/8080715.gif";
leftlinksuspent="http://goto.sohu.com/goto.php3?code=benq-sh176-alumniclass80";
leftflashsuspent="http://images.sohu.com/cs/button/haifeisi/80800707.swf?clickthru=http://goto.sohu.com/goto.php3?code=haifeisi-g32-banjifly";
}
else
{
leftformat="";
}
leftmargin1="15";
topleft="90";


if (monthss==6 & (days==19 | days==20 | days==21 | days==22))
{
rightformat="flash";
rightpicsuspent="http://images.sohu.com/cs/button/wahaha/8080513.gif";
rightlinksuspent="http://goto.sohu.com/goto.php3?code=wahaha-sh231-alumniclass8080";
rightflashsuspent="http://images.sohu.com/cs/button/philips/8080720.swf?clickthru=http://goto.sohu.com/goto.php3?code=philips-sh185-alumniclass80";
}
else if (monthss==6 & (days==10 | days==11 | days==17 | days==18))
{
rightformat="gif";
rightpicsuspent="http://images.sohu.com/cs/button/shida/8080.gif";
rightlinksuspent="http://goto.sohu.com/goto.php3?code=shida-g51-alumnibanjifly";
rightflashsuspent="http://images.sohu.com/cs/button/haitian/800603.swf?clickthru=http://goto.sohu.com/goto.php3?code=haitian-bj328alumniclassfly";
}
else
{
rightformat="";
}

rightmargin1="550";
topright="90";
flytransparency="";
liumeitiformat="";
lmtleft="430";
lmttop="220";
lmtwidth="200";
lmtheight="150";
lmtransparency="";
liumeiti="http://images.sohu.com/cs/button/feiyate/2001500926.swf";
liumeititime="5000";


document.write("<script language=javascript src=http://images.sohu.com/cs/jsfile/allfly.js></SCRIPT>");