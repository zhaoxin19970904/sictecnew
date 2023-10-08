var current = null;		// indicating current page
function expand(id) {
	var target = document.getElementById(id);
	if (!target)
		return;
	if (current) {
		current.style.display = "none";
	}
	current=target;
	target.style.display="block";
}

/* 
 * This is called from the right window to update the folder list
 * The input is the fragment of html to be inserted to tablist.z
 */
function updateFolderList(s) {
	var target = document.getElementById('tablist.z');
	if (target) {
		target.innerHTML = s;
	} else
		alert ('cannot find tablist.z');
}

menuopen=-1;
menuurlprefix='';

function findObj(n) {
	var x;
	
	d=document;
	if(!(x=d[n])&&d.all) x=d.all[n];
	if(!x && d.getElementById) x=d.getElementById(n);
	return x;
}

function hidemenu(menuid) {
	menuobj=findObj("menu"+menuid)
	menuheadobj=findObj("menuhead"+menuid)
	if(!menuheadobj)return;
	if (menuobj) {
		menuobj.style.visibility='hidden';
	}
	menuheadobj.className='menuhead';

	menuopen=-1;
}

function showmenu(menuid) {
	menuobj=findObj("menu"+menuid);
	menuheadobj=findObj("menuhead"+menuid)
	if(!menuheadobj)return;
	if (menuobj) {
		dh=document.body.clientHeight;
		menupos=menuheadobj.parentElement.offsetTop-document.body.scrollTop;
		menuheight=menuobj.clientHeight;
		if((menupos+menuheight>dh)&&(menupos>0)) {
			menutop=dh-(menuheight+menupos)-2;
			if (menutop<-menupos) menutop=-menupos;
			menuobj.style.top=menutop;
		} else {
			menuobj.style.top=0;
		}
		
		menuobj.style.visibility='visible';
		//menuobj.filters.opacity=55;
		//menutobj.filters.opacity=55;

		menuheadobj.className='menuhead_opensub';
	} else {
		menuheadobj.className='menuhead_open';
	}

	menuopen=menuid;
}

function menuhead_rollon() {
	o=window.event.srcElement;
	openmenuid=o.id.substr(8);

	if(menuopen>=0) hidemenu(menuopen);

	showmenu(openmenuid);
}

function menu_rollon() {
	o=window.event.srcElement;
	if (o.className == "menuopt") o.className = "menuopt_over";
}

function menu_rolloff() {
	o=window.event.srcElement;
  	if (o.className.substring(0,7) == "menuopt") o.className = "menuopt";
}

function menu_mousedown() {
	o=window.event.srcElement;
	if (o.className == "menuopt_over") o.className = "menuopt_click";
 	if (o.className == "menuhead_open") o.className = "menuhead_click";
 	if (o.className == "menuhead_opensub") o.className = "menuhead_clicksub";
}

function menu_mouseup() {
	o=window.event.srcElement;
 	if (o.className == "menuopt_click") o.className = "menuopt_over";
 	if (o.className == "menuhead_click") o.className = "menuhead_open";
 	if (o.className == "menuhead_clicksub") o.className = "menuhead_opensub";
}

function menu_mouseclick() {
	o=window.event.srcElement;
	if (o.className == "menuopt_over") {
		menusubid=o.id.split(".");
		menuid=menusubid[0];
		menuoptid=menusubid[1];
		t=menuoptions[menuid][menuoptid].split('|');
		newurl=t[1];
		if (newurl.substr(0,7)!='http://') newurl=menuurlprefix + newurl;
		if (newurl.substr(0,3)=='js:') 
		eval(newurl.substr(3,newurl.length));
		else
		window.location=newurl;
	} else if ((o.className == "menuhead_open")||(o.className == "menuhead_opensub")) {
		menuid=o.id.substr(8);
		t=menuoptions[menuid][0].split('|');
		newurl=t[1];
		if (newurl.substr(0,7)!='http://') newurl=menuurlprefix + newurl;
		if (newurl.substr(0,3)=='js:') 
		eval(newurl.substr(3,newurl.length));
		else
		window.location=newurl;
	}
}

function document_mouseover() {
	if(menuopen>=0) {
		o=window.event.srcElement;
		menuobj=findObj("menu"+menuopen);
		if (menuobj) {
			if ((menuobj.style.visibility == "visible")&&(o.id.length==0)) hidemenu(menuopen);
		} else {
			menuobj=findObj("menuhead"+menuopen);
			if ((menuobj.className == "menuhead_open")&&(o.id.length==0)) menuobj.className="menuhead";
		}
	}
}

function menu_build() {
	for (x in menuoptions) {
		subopts=menuoptions[x].length;
		vpos=250+(x*17);
		for (y in menuoptions[x]) {
			t=menuoptions[x][y].split('|');
			if(y==0) {
				cc=t[2];
				document.write ('<span class="menuspan" style="position: absolute; top: ' + vpos + ';">&nbsp;&nbsp;');
				if(subopts>1){
				document.writeln ('<span class="menuhead" id="menuhead'+x+'" unselectable="on">'+t[0]+'</a></span>');
				}else{
					if(cc!=200){					
					document.writeln ('<span class="menuheadx" id="menuhead'+x+'" unselectable="on" style="color:#000000">&middot;'+t[0]+'</span>');
					}else{
					document.writeln ('<span class="menuhead" id="menuhead'+x+'" unselectable="on">'+t[0]+'</span>');
					}
				}				
				if(subopts>1) document.writeln ('<span id="menu'+x+'" class="menu"><table width="'+t[2]+'" border="0" cellspacing="0" cellpadding="0">');
			} else {
				document.writeln ('<tr><td class="menuopt" id="'+x+'.'+y+'" unselectable="on">'+t[0]+'</td></tr>');
			}
		}
		if(subopts>1) document.writeln('</table></span>');
		document.write('</span>');
	}
}

function menu_addevents() {
	for (x in menuoptions) {
		className = eval('menuhead'+x+'.className');
		if(className == 'menuhead')eval('menuhead'+x+'.onmouseover = menuhead_rollon');
		if(className == 'menuhead')eval('menuhead'+x+'.onmousedown = menu_mousedown');
		if(className == 'menuhead')eval('menuhead'+x+'.onmouseup = menu_mouseup');
		if(className == 'menuhead')eval('menuhead'+x+'.onclick = menu_mouseclick');
		if (menuoptions[x].length>1) {
			className = eval('menu'+x+'.className');
			if(className == 'menu')eval('menu'+x+'.onmouseout = menu_rolloff');
			if(className == 'menu')eval('menu'+x+'.onmouseover = menu_rollon');
			if(className == 'menu')eval('menu'+x+'.onmousedown = menu_mousedown');
			if(className == 'menu')eval('menu'+x+'.onmouseup = menu_mouseup');
			if(className == 'menu')eval('menu'+x+'.onclick = menu_mouseclick');
		}
	}
	document.onmouseover=document_mouseover;
}

