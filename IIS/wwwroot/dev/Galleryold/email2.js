var back="",prefix="",suffix = "",pos, count = 0, content = new Array("","");
var defaultstyle = "<link rel=\"stylesheet\" type=\"text/css\" href=\"gallery.css\">";
if(document.images)
{var Over = new Image();
var Out = new Image();
Over.src="images/exit_over.gif";
Out.src="images/exit_out.gif";
}else{}

function setprefix(pre)
{prefix = pre;}
function setsuffix(suf)
{suffix = suf;}

function email_page()
{
//parent.right.document.writeln ("");
parent.right.document.writeln ("<html><head><title>New Page 3</title></head>" + defaultstyle + "<body   background=\"images/creamGradient.jpg\" style=\"background-repeat:repeat-x; background-attachment:fixed\" topmargin=\"3\" leftmargin=\"3\" rightmargin=\"3\" bottommargin=\"3\">");
parent.right.document.writeln ("<p align=\"center\"><b>Email Content</b></p><a href='javascript:parent.e_mail();'>Pre-View</a>");
parent.right.document.writeln ("<form method=\"POST\" action=\"--WEBBOT-SELF--\">");
parent.right.document.writeln ("<!--webbot bot=\"SaveResults\" U-File=\"fpweb:///_private/form_results.csv\" S-Format=\"TEXT/CSV\" S-Label-Fields=\"TRUE\" -->");
parent.right.document.writeln ("<input type=\"submit\" value=\"Veiw Email\" name=\"VeiwEmail\" style=\"float: right\"></form></body></html>");
parent.right.document.close ();
}

function preview()
{
str = ""
for(i;i<count;i+=2)
{
str+="&i=" + content[pos];
}
window.open ("create_email.asp?"+str);
}

function page()
{
parent.rightB.document.clear();
parent.rightB.document.write("<html><head><title>Content</title></head>" + defaultstyle);
parent.rightB.document.write("<body   background=\"images/creamGradient.jpg\" style=\"background-repeat:repeat-x; background-attachment:fixed\" topmargin=\"0\" leftmargin=\"4\" rightmargin=\"4\" bottommargin=\"0\">");
parent.rightB.document.write("<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"115\">");
Refresh();
parent.rightB.document.write("</Table></body></html>");
parent.rightB.document.close();
parent.rightB.focus ();
}



function Refresh()
{
	max = count;
	count = 0;
	for(pos = 0;pos < max;pos+=2)
	{
		if(content[pos] != "")
		{AddRow(content[pos],content[pos + 1]  );
		content[count] = content[pos];
		content[count+1] = content[pos +1];
		count+=2;}
	}
		 
}
function RemoveImage(Ipos)
{
	content[Ipos] = "";
	content[Ipos + 1] = "";
	page();
}

function AddRow(image,code)
{
	parent.rightB.document.writeln("<tr><td align='right' valign='bottom' height='80px'>");
	parent.rightB.document.writeln("<img width='115px' height='80px' src='" + prefix + image + suffix + "'>");
	parent.rightB.document.writeln(code + "<img onclick=\"parent.RemoveImage(" + count + ");\" onmouseout=\"this.src=parent.Out.src;\" onmouseover=\"this.src=parent.Over.src;\" border=\"0\" src=\"images/exit_out.gif\" width=\"26\" height=\"25\" style=\"cursor:hand;position: relative; z-index: 1; top:-80px;\"></td></tr>");
	parent.rightB.document.writeln("</td></tr>");
}

function AddImage(image,code)
{
	content[count] = image;
	content[count+1] = code;
	count+=2;
	page();
}

function e_mail()
{
	str="";

	for(pos = 0; pos < count;pos+=2)
	{
		str+= "&I=" + content[pos];
	}
window.open ("creat_email.asp?I=-1" + str, "", "toolbar=yes,menubar," + "width=800,height=600,Left=" + ((screen.Width / 2) - (800 / 2)) + ",Top=" + ((screen.Height / 2) - (600 / 2)) + ",scrollbars=yes");
//document.location.replace("creat_email.asp?I=-1" + str);
}

function pre_view()
{
	win = loadWin(600,800,"Pre - Email Options",false,false,false);
	win.document.writeln(defaultstyle);
	win.document.writeln("<frameset framespacing='0' frameborder='0' rows='270px,*'><frameset framespacing='0' frameborder='0' cols='300px,*'><frame name='top' src='email_options.htm' scrolling='no'><frame name='TLayout' scrolling='no' marginwidth='0' marginheight='0'></frameset>");
	
	//win.document.writeln("<body topmargin=\"3\" leftmargin=\"3\" rightmargin=\"3\" bottommargin=\"3\"><script language=\"javascript\">\nfunction b(){alert('ssssss');}\n</script><iframe src=\"email_options.htm\" width=\"100%\" height=\"100%\">browser</iframe>");
	
	finWin(win);
}

function loadWin ( w, h, t, s, tb,mb)
{
if(s){sc = ",scrollbars=yes"}else{sc = ",scrollbars=no"}
Tbar="";if(tb){Tbar="toolbar=yes,"}
if(mb){Tbar+="menubar=yes,"}
LoadWin = window.open ("", "", Tbar + "width=" + w + ",height=" + h + ",Left=" + ((screen.Width / 2) - (w / 2)) + ",Top=" + ((screen.Height / 2) - (h / 2)) + sc);
LoadWin.document.write ("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=windows-1252\"><Title>" + t + "</Title>");
return LoadWin;
}
function finWin(window)
{
	window.document.write ("</Body></Html>");
	window.document.close();
}

function preTLayout()
{
alert('sdsd');
	win.document.all.TLayout.document.writeln ("<html><head><title>New Page 3</title></head><body>");
	//win.parent.TLayout.close();
alert('sdsd2');
	win.TLayout.document.writeln ("<table border=\""+ window.Border +"\" cellpadding=\"0\" cellspacing=\"0\" width=\"201\" height=\"74\" bordercolor=\"#FF00FF\" style=\"border-collapse: collapse\">");
	win.TLayout.document.writeln ("<tr><td height=\"36\">&nbsp;</td><tdwidth=\"95\">&nbsp;</td></tr><tr><td>&nbsp;</td><td>&nbsp;</td></tr></table>");

}

function WriteMail(border,collasped,spacing,BCR,BCG,BCB,PCR,PCG,PCB,coloms,baseW,Images)
{

	parent.temp.document.writeln("<form target='_blank' name=\"TLayout\" method=\"POST\" action=\"create_email.asp?Email=" + Date() + "\">");
	
	parent.temp.document.writeln("<input type='text' name='Count' value='" + count + "'>");
	parent.temp.document.writeln("<input type='text' name='Border' value='" + border + "'>");
	parent.temp.document.writeln("<input type='text' name='Collasped' value='" + collasped + "'>");
	parent.temp.document.writeln("<input type='text' name='Spacing' value='" + spacing + "'>");
	parent.temp.document.writeln("<input type='text' name='PageColor' value='RBG(" + (PCR +","+ PCG +","+ PCB) + ")'>");
	parent.temp.document.writeln("<input type='text' name='BorderColor' value='RBG(" + (BCR +","+ BCG +","+ BCB) +")'>");
	
	parent.temp.document.writeln("<input type='text' name='Coloms' value='" + coloms + "'>");
	parent.temp.document.writeln("<input type='text' name='BaseWidth' value='" + baseW + "'>");

	
	for(c=0,info=0;c<count;c++,info=info+2)
	{
		parent.temp.document.writeln("<input type='text' name='Iurl' value='" + content[c] + "'>");
	}
	
	parent.temp.document.writeln("</form>");
	parent.temp.document.close();
	parent.temp.document.TLayout.submit();

/*
	url="create_email.asp?";
	url+="Count="+count;
	url+="&Border="+border;
	url+="&Collasped="+collasped;
	url+="&Spacing="+spacing;
	url+="&PageColor=RBG(" + (PCR +","+ PCG +","+ PCB) +")";
	url+="&BorderColor=RBG(" + (BCR +","+ BCG +","+ BCB) +")";
	url+="&Coloms="+coloms;
	url+="&BaseWidth="+baseW;
	for(c=0,info=0;c<count;c++,info=info+2)
	{
	 url+="&Iurl="+ content[c] + "&Quality=30&Width=400"// + Images[info +1] + "&Width=" + Images[info];
	}

	window.open(url,"","channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=1,resizable=1,scrollbars=1,status=0,toolbar=1");
*/
}

function WriteForm()
{
	for(pos = 0; pos < count;pos++)
	{
		document.writeln("<input type='text' name='image' value='" + content[pos] +"'>");
	}
}