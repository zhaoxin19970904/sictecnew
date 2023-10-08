<%
function UBBCode(strContent)
if not isnull(strContent) then
  if cint(bbs_seting(2))=1 then
  strContent=dvHTMLEncode(strContent)
  end if
  if cint(bbs_seting(3))=0 then
  UBBCode=strContent
  exit function
  end if

dim re
dim po,ii
dim reContent
Set re=new RegExp
re.IgnoreCase =true
re.Global=True

re.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")
re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")

re.Pattern="(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"">$2</A>")
re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")
	re.Pattern="\[DIR=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/DIR]"
	strContent=re.Replace(strContent,"<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>")
	re.Pattern="\[QT=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/QT]"
	strContent=re.Replace(strContent,"<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>")
	re.Pattern="\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
	strContent=re.Replace(strContent,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3  width=$1 height=$2></embed></object>")
	re.Pattern="\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
	strContent=re.Replace(strContent,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")

	re.Pattern="(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<a href=""$2"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")

	re.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(.[^\[]*)(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<a href=""$4"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4""><PARAM NAME=quality VALUE=high><embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")

re.Pattern="\[color=(.[^\[]*)\](.[^\[]*)\[\/color\]"
strContent=re.Replace(strContent,"<font color=$1>$2</font>")
re.Pattern="\[face=(.[^\[]*)\](.[^\[]*)\[\/face\]"
strContent=re.Replace(strContent,"<font face=$1>$2</font>")
re.Pattern="\[align=(center|left|right)\](.*)\[\/align\]"
strContent=re.Replace(strContent,"<div align=$1>$2</div>")

    re.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
    strContent=re.Replace(strContent,"<table width=99% bgcolor=#0000FF cellspacing=1 border=0 cellpadding=0 align=center><tr> <td><table width=100% border=0 bgcolor=#FFFFFF cellpadding=0 cellspacing=3 align=center> <tr><td >$2</td> </tr></table> </td></tr></table><br>")
    re.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
    strContent=re.Replace(strContent,"<table width=99% bgcolor=#0000FF cellspacing=1 border=0 cellpadding=0 align=center><tr> <td><table width=100% border=0 bgcolor=#FFFFFF cellpadding=0 cellspacing=3 align=center> <tr><td >$2</td> </tr></table> </td></tr></table><br>")

re.Pattern="\[fly\](.*)\[\/fly\]"
strContent=re.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
re.Pattern="\[move\](.*)\[\/move\]"
strContent=re.Replace(strContent,"<MARQUEE scrollamount=3>$1</marquee>")	
re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
strContent=re.Replace(strContent,"<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
strContent=re.Replace(strContent,"<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")
re.Pattern="(\[sound\])(.[^\[]*)(\[\/sound\])" 
strContent=re.Replace(strContent,"<a href=""$2"" target=_blank><IMG SRC=pic/mid.gif border=0 alt='背景音乐'></a><bgsound src=""$2"" loop=""-1"">")


re.Pattern="\[i\](.[^\[]*)\[\/i\]"
strContent=re.Replace(strContent,"<i>$1</i>")
re.Pattern="\[u\](.[^\[]*)(\[\/u\])"
strContent=re.Replace(strContent,"<u>$1</u>")
re.Pattern="\[b\](.[^\[]*)(\[\/b\])"
strContent=re.Replace(strContent,"<b>$1</b>")
re.Pattern="\[size=([1-4])\](.[^\[]*)\[\/size\]"
strContent=re.Replace(strContent,"<font size=$1>$2</font>")
re.Pattern="(\[em(.[^\[]*)\])"
strContent=re.Replace(strContent,"<img src=pic/em$2.gif border=0 align=middle>")

re.Pattern="(\[IMG\])(.[^\[]*)(\[\/IMG\])"
strContent=re.Replace(strContent,"<a href=""$2"" target=_blank><IMG SRC=""$2"" border=0 onLoad=""javascript:if(this.width>580)this.width=580"" alt=按此在新窗口浏览图片></a><br>")


'自动识别网址
re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/urlIE.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)$"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/urlIE.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
strContent = re.Replace(strContent,"$1<img align=absmiddle src=pic/urlIE.gif border=0><a target=_blank href=$2>$2</a>")

'自动识别www等开头的网址
re.Pattern = "([^(http://|http:\\)])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/urlIE.gif border=0><a target=_blank href=http://$2>$2</a>")

strContent=ChkBadWords(strContent)
UBBCode=replace(strContent,vbCrlf,"<br>")
end if
end function

function ChkBadWords(fString)
    dim bwords,ii,rs
    if not(isnull(BadWords) or isnull(fString)) then
    bwords = split(BadWords, "|")
    for ii = 0 to ubound(bwords)
        fString = Replace(fString, bwords(ii), string(len(bwords(ii)),"*")) 
    next
    ChkBadWords = fString
    end if
end function

function userubbqm(strContent)
if not isnull(strContent) then
dim re
dim po,ii
dim reContent
Set re=new RegExp
re.IgnoreCase =true
re.Global=True

re.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")
re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")

re.Pattern="(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"">$2</A>")
re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")

	re.Pattern="(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<a href=""$2"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")

	re.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(.[^\[]*)(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<a href=""$4"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4""><PARAM NAME=quality VALUE=high><embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")

re.Pattern="\[color=(.[^\[]*)\](.[^\[]*)\[\/color\]"
strContent=re.Replace(strContent,"<font color=$1>$2</font>")
re.Pattern="\[face=(.[^\[]*)\](.[^\[]*)\[\/face\]"
strContent=re.Replace(strContent,"<font face=$1>$2</font>")
re.Pattern="\[align=(center|left|right)\](.*)\[\/align\]"
strContent=re.Replace(strContent,"<div align=$1>$2</div>")

re.Pattern="\[fly\](.*)\[\/fly\]"
strContent=re.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
re.Pattern="\[move\](.*)\[\/move\]"
strContent=re.Replace(strContent,"<MARQUEE scrollamount=3>$1</marquee>")	
re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
strContent=re.Replace(strContent,"<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
strContent=re.Replace(strContent,"<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")
re.Pattern="(\[sound\])(.[^\[]*)(\[\/sound\])" 
strContent=re.Replace(strContent,"<a href=""$2"" target=_blank><IMG SRC=pic/mid.gif border=0 alt='背景音乐'></a><bgsound src=""$2"" loop=""-1"">")

re.Pattern="\[i\](.[^\[]*)\[\/i\]"
strContent=re.Replace(strContent,"<i>$1</i>")
re.Pattern="\[u\](.[^\[]*)(\[\/u\])"
strContent=re.Replace(strContent,"<u>$1</u>")
re.Pattern="\[b\](.[^\[]*)(\[\/b\])"
strContent=re.Replace(strContent,"<b>$1</b>")
re.Pattern="\[size=([1-4])\](.[^\[]*)\[\/size\]"
strContent=re.Replace(strContent,"<font size=$1>$2</font>")
re.Pattern="(\[em(.[^\[]*)\])"
strContent=re.Replace(strContent,"<img src=pic/em$2.gif border=0 align=middle>")

re.Pattern="(\[IMG\])(.[^\[]*)(\[\/IMG\])"
strContent=re.Replace(strContent,"<a href=""$2"" target=_blank><IMG SRC=""$2"" border=0 onLoad=""javascript:if(this.width>580)this.width=580"" alt=按此在新窗口浏览图片></a><br>")

strContent=ChkBadWords(strContent)
userubbqm=replace(strContent,vbCrlf,"<br>")
end if
end function

function ChkBadWords(fString)
    dim bwords,ii,rs
    if not(isnull(BadWords) or isnull(fString)) then
    bwords = split(BadWords, "|")
    for ii = 0 to ubound(bwords)
        fString = Replace(fString, bwords(ii), string(len(bwords(ii)),"*")) 
    next
    ChkBadWords = fString
	end if
end function

function dvHTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "<I></I>&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    fString=ChkBadWords(fString)
    dvHTMLEncode = fString
end if
end function

function dispscript(strContent)
dim re
dim po,ii
dim reContent
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
re.Pattern="(javascript)"
strContent=re.Replace(strContent,"&#106avascript")
re.Pattern="(jscript:)"
strContent=re.Replace(strContent,"&#106script:")
re.Pattern="(js:)"
strContent=re.Replace(strContent,"&#106s:")
re.Pattern="(value)"
strContent=re.Replace(strContent,"&#118alue")
re.Pattern="(about:)"
strContent=re.Replace(strContent,"about&#58")
re.Pattern="(file:)"
strContent=re.Replace(strContent,"file&#58")
re.Pattern="(document.cookie)"
strContent=re.Replace(strContent,"documents&#46cookie")
re.Pattern="(vbscript:)"
strContent=re.Replace(strContent,"&#118bscript:")
re.Pattern="(vbs:)"
strContent=re.Replace(strContent,"&#118bs:")
re.Pattern="(on(mouse|exit|error|click|key))"
strContent=re.Replace(strContent,"&#111n$2")
end function
%>
