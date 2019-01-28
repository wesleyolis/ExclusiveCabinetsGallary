<%@ LANGUAGE = VBScript %>
<!--METADATA TYPE="typelib" 
uuid="00000206-0000-0010-8000-00AA006D2EA4" -->
<% Session.Timeout = 20 %>
<%
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->

<%
Response.Buffer = true
Dim dir,udir,str,email,admin, rs,rsi,FilRange, FIL,Group1,Group2, Range,Description, Page

Page = 1
Group1 =""
Group2 =""
Range = ""
Description = ""
FIL=""
FilRange = -1
str=""
If Session("Admin") Then
Session("Admin") = true 
admin=true
'str = str & "&admin"
Else
Session("Admin") = false
admin=false
End IF
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

if Request.QueryString("Range").Count > 0 Then
Range = Request.QueryString("Range")
IF (Not Range = "-1") AND Range >= 0 And Range < 65536 Then
FIL = FIL &  " AND (([Images.Range])=" & Range & ") "
End IF
End IF

if Request.QueryString("Des").Count > 0 Then
Description = Request.QueryString("Des")
IF  (Not Description = "-1") AND  Description >= 0 And Description < 65536 Then
FIL = FIL &  " AND (([Images.Description])=" & Description & ") "
End IF
End IF

if Request.QueryString("Color").Count > 0 Then
Color = Request.QueryString("Color")
IF  (Not Color = "-1") AND  Color >= 0 And Color < 65536 Then
FIL = FIL &  " AND (([Images.Color])=" & Color & ") "
END IF
End IF

if Request.QueryString("Ulti")<> "" Then
Ulti = Request.QueryString("Ulti")
FIL = FIL & " AND (InStr(UCase(([Images.Code])), UCase('" & Ulti & "')) > 0)"       '  " AND (UCase(([Images.Code]))=UCase('" & Ulti & "')) "    (InStr(UCase(([Images.Code])), UCase('" & Ulti & "')) > 0)
End IF

if Request.QueryString("Client") <> "" Then
Client = Request.QueryString("Client")
FIL = FIL & " AND (InStr(UCase(([Images.Client])), UCase('" & Client & "')) > 0)"  '' " AND (UCase(([Images.Client]))=UCase('" & Client & "')) "
End IF


if Request.QueryString("Email").Count > 0 Then
email = true
str = str & "&email=true"
End IF

if Request.QueryString("Page") <> "" Then
Page = Request.QueryString("Page")
Else
Page = 1
End IF



if Request.QueryString("d").Count = 0 Then
dir = -2
else

if (Request.QueryString("d")>= -1) And (Request.QueryString("d") < 655536) Then
dir = Request.QueryString("d")

End IF
If Not dir = "-2" Then
FIL = FIL & "AND(([Images.Dir]= "&dir&"))"


if Request.QueryString("u").Count <> 0 And (Request.QueryString("u")>= -1) And (Request.QueryString("u") < 655536) Then
udir = Request.QueryString("u")
Else
Set rs = Conn.Execute("SELECT Dir.[UDir] FROM Dir where ([Dir.Dir]= " & dir & ");")
If Not rs.Eof Then
udir = rs(0)
End IF
rs.Close
rs = NULL

End IF
END IF
End IF
If Not dir ="-2" Then

sql = "SELECT Dir.Dir, Dir.Name, Dir.Description, Dir.[UDir] FROM Dir where ([Dir.UDir]= " & dir & ");"
Set rs = Conn.Execute(sql)
'set rs = Server.CreateObject("ADODB.recordset")
'rs.Open sql, conn
END IF

'Response.write(FIL)
sql = "SELECT Images.Image, Images.Code, Images.Client, Ranges.Name, Descriptions.Name, Colors.Name, Images.Width, Images.Height, Images.Depth,IIF([Images.Price]=0,'P.O.A','R ' & [Images.Price]) As Price,Images.Edge, Images.Info"_
&" FROM Descriptions RIGHT JOIN (Colors RIGHT JOIN (Ranges RIGHT JOIN Images ON Ranges.Range = Images.Range) ON Colors.Color = Images.Color) ON Descriptions.Description = Images.Description"_
&" WHERE ((1=1) " & FIL & " ) ORDER BY Descriptions.Name,Ranges.Name, Colors.Name;"

set rsi = Server.CreateObject("ADODB.recordset")
rsi.Open sql, conn, adOpenStatic
rsi.PageSize = 30
If (Page > rsi.PageCount) And (Page < 0) Then
Page = 1
End IF
If Not rsi.Eof Then
rsi.AbsolutePage = Page
End IF
%>

<html>
<head>
<title>Exclusive Cabinets Gallary</title>
</head>
<link rel="stylesheet" href="gallery.css" type="text/css">
<style>
 .udot {border-bottom-style: dashed; border-bottom-width: 1px; padding: 2px; }
 .t1{font-size:11pt;border: 2px solid #cc907e; padding: 0}
 .GHeading{border-bottom: 2px solid #FF0000; font-size:14pt; color:#DE4010; font-weight:bold}
</style>
<script language="javascript">
<!--
function d(img)
{
if(document.getElementById("i"+ img).style.display == 'block')
document.getElementById("i"+ img).style.display = 'none';
else
document.getElementById("i"+ img).style.display = 'block';
}
<%if email Then%>	
parent.setprefix("thumbs/");
parent.setsuffix(".jpg");
<%End IF%>

<%if admin Then%>	

function all_del()
{
	var answer = confirm ("Delete this Entry?") ;
	if (answer)
	{
		var answer2 = confirm ("Delete the Original Image as well") ;
		if(answer2){
		document.all.orig.value = '1';
				}
		else{
		document.all.orig.value = '0';
			}
		document.Del_ALL.submit()
	}
}


function edit(img)
{
document.location.replace("ImagesAdd.asp?d=<%=dir%>&Key="+img);
}
function addsub(img)
{
document.location.replace("Images_SubAdd.asp?I="+img);
}


function image(img)
{
document.location.replace("ImagesAdd2.asp?d=<%=dir%>&Key="+img);
}


function del(img)
{
var answer = confirm ("Delete this Entry?") 
if (answer)
{
	var answer2 = confirm ("Delete the Original Image as well") 
	if(answer2){
	document.location.replace("Imagesdelete.asp?d=<%=dir%>&Key="+img+"&o=");}
	else{
	document.location.replace("Imagesdelete.asp?d=<%=dir%>&Key="+img);}
}
}
<% If Not dir = "-2" Then%> 
function DirDel()
{
var answer2 = confirm ("Are you sure that you want to del this directory!") 
	if(answer2){
document.location.replace('Dirdelete.asp?d=<%=udir%>&Key=<%=dir%>');
}
}
<%END IF%>
<%End IF%>
//-->
</script>
<body background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed">
<%if Session("Admin") = True THEN%>
<Div align="right"><Form action="login.asp" method="post">
<%If Request.QueryString("Email") = "true" Then%><input  type="hidden" name="Email" value="true"><%END IF%>
<input type="submit" name="Action" value="Logout">
</Form></Div>
<%END IF%>
<%if udir <> "" Then%><a href="dirbrowser.asp?d=<%=udir%><%=str%>">Back to Parent</a><%End IF%>
<table border="0" cellpadding="0" cellspacing="0" id="table1" width="800">
<% If Not dir = "-2" Then %>
<%if admin Then%>
<tr><td class="box2" height="32" colspan="2">		
<b>Directory</b>		
<input class="text" onclick="document.location.replace('Diradd.asp?d=<%=dir%>');" type="button" value="Add" name="add_d">&nbsp; 
<%If udir <> ""  Then%>
<input class="text" onclick="document.location.replace('Diradd.asp?key=<%=dir%>');" type="button" value="Edit" name="edit_d">
<input class="text" onclick="<%if (Not rs.eof) Or (Not rsi.Eof) Then%>alert('This Directory is not empty.\n Can not Delete!');<%Else%>DirDel();<%End If%>" type="button" value="Delete" name="delete_d">&nbsp;&nbsp;&nbsp;
 <%End If%>
<tr><td colspan="2">
<%End If%>
<tr><td colspan="2">
<%Do While Not rs.eof%>
<a href="dirbrowser.asp?d=<%=rs.Fields("dir")%>&u=<%=dir%><%=str%>"><%=rs.Fields("Name")%></a>&nbsp;;&nbsp;
<%
rs.MoveNext
Loop
rs.Close()
%>
</td></tr>
<%END IF%>


<tr><td class="box2" height="32" colspan="2">	
<script language="javascript">
<!--

function filter()
{
document.location.replace("Dirbrowser.asp?" + CreateLink());
}

function page(p)
{
document.location.replace("Dirbrowser.asp?" + CreateLink() + "&Page=" + p);
}


function CreateLink()
{
str="";
<%If dir <> 2 Then%>
str+="d=<%=dir%>";
<%END IF%>
v = document.all.x_Range.value;
if(v != -1)
{str+="&Range=" + v}

v = document.all.x_Description.value;
if(v != -1)
{str+="&Des=" + v}

v = document.all.x_Color.value;
if(v != -1)
{str+="&Color=" + v}

v = document.all.Ulti.value;
if(v != "")
{str+="&Ulti=" + v}

v = document.all.Client.value;
if(v != "")
{str+="&Client=" + v}


<%if email Then%>
str+="&Email=true";
<%END IF%>

	return str;
}
//-->
</script>
<b>Range</b>&nbsp;
<%
Set rs = conn.Execute("SELECT Ranges.Range, Ranges.Name FROM Ranges ORDER BY Ranges.Name;")
%>
<select onchange="filter()" class="input" id="x_Range" name="x_Range" size="1" >
<option value="-1">No Filter</option>
<%
j = 0
Do While (Not rs.Eof)
val = rs.Fields(0)
If val & "" = Range Then
%><option  selected value="<%=val%>"><%=rs.Fields(1)%></option>
<%Else%>
<option value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<%
End IF
rs.MoveNext
Loop
Response.write("</select>")
%>

 
</select> <b>&nbsp;Description&nbsp; </b><%
Set rs = conn.Execute("SELECT Descriptions.Description, Descriptions.Name FROM Descriptions ORDER BY Descriptions.Name;")
%>
<select  onchange="filter();"  class="input" id="x_Description" name="x_Description" size="1" >

<option value="-1">No Filter</option>
<%
Do While Not rs.Eof
val = rs.Fields(0)
If val & ""  = Description  Then
%><option selected value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<%Else%>
<option value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<%
End IF
rs.MoveNext
Loop
Response.write("</select>")
%>

<%
Set rs = conn.Execute("SELECT Colors.Color, Colors.Name FROM Colors ORDER BY Colors.Name;")
%>

</select><b>&nbsp; Colour&nbsp; </b>&nbsp;<select  onchange="filter()"  class="input" id="x_Color" name="x_Color" size="1" >
<option value="-1">No Filter</option>
<%
Do While Not rs.Eof
If rs.Fields(0) & "" = Color Then
%><option selected value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<%Else%>
<option value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<%
End IF
rs.MoveNext
Loop%>
</select>
</td></tr>
<tr>
	<td class="box2" height="32px" width="505">
<b>Ulti Sales </b>&nbsp; <input type="text" name="Ulti" id="Ulti" size="16" maxlength="16" value="<%=Request.QueryString("Ulti")%>">&nbsp;&nbsp;
<b>Client Ref </b>&nbsp; <input type="text" name="Client" id="Client" size="16" maxlength="16"  value="<%=Request.QueryString("Client")%>">&nbsp;&nbsp; <input type="button" onclick="filter();" value="Search">
</td>
	<td class="box2" height="32px" width="295">
<b>Pages:&nbsp;&nbsp;<%
Page = rsi.AbsolutePage
p=1
If (Page -5)>0 Then
p = Page -5
End IF
pend = rsi.PageCount
If (Page + 5) <= rsi.PageCount Then
pend = Page + 5
End If
While (p <= rsi.PageCount) And p <=pend
If p = Page Then
%><font size="3"><%=p%></font>&nbsp;&nbsp;
<%Else
%><a href="javascript:page(<%=p%>);"><font size="3"><%=p%></font></a>&nbsp;&nbsp;<%
End IF
p=p+1
Wend%></b>
</td>
</tr>
<tr><td align="left" colspan="2">
<%if admin Then%>
<form method="POST" action="Imagesdelete.asp?d=<% If Not dir = "-2" Then %><%=dir%><%ELSE%>-1<%END IF%>" name="Del_ALL">
<input type="hidden" name="orig" value="1">
<div align="left">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr><td class="box2" height="32">	
<b>Entries</b>
<input onclick="document.location.replace('Imagesadd.asp?d=<% If Not dir = "-2" Then %><%=dir%><%ELSE%>-1<%END IF%>');" class="text" type="button" value="Add Image" name="Add_Image">
<input onclick="all_del()" class="text" type="button" value="Delete All" name="delete_all">&nbsp; 
</td></tr>
<tr><td>
<p align="center">
<%End If%>
<%If rsi.eof Then
%><br><font color="#DE4010" size="4"><b>No Results Found</b></font>
<%
End IF
Rows = 2
RCount = 0
%>
</p>
<table border="0" cellpadding="0" cellspacing="20" width="100%">
<%
Do While Not rsi.eof And (j < rsi.PageSize )
j = j + 1
change = False
If (Range = "") AND (Not Group1 = rsi.Fields("Ranges.Name")) Then
Group1 = rsi.Fields("Ranges.Name")
change = True
End IF

If (Description ="") AND (Not Group2 = rsi.Fields("Descriptions.Name")) Then
Group2 = rsi.Fields("Descriptions.Name")
change = True
End IF

If change = True Then

If (Range = "" Or Description ="") And Rcount > 0 And Rcount < Rows Then
RCount = 0
%><td colspan="<%=(Rows-Rcount)%>">&nbsp;</td>
<% End IF

If (Range = "" And Description ="") Then
%>
<tr><td colspan="<%=Rows%>" class="Gheading"><%=Group1%>&nbsp;-&nbsp;<%=Group2%>&nbsp;</td></tr>
<%
Else
If (Range = "") Then
%><tr><td colspan="<%=Rows%>" class="Gheading"><%=Group1%>&nbsp;</td></tr>
<%
Else%>
<tr><td colspan="<%=Rows%>" class="Gheading"><%=Group2%>&nbsp;</td></tr>
<%
End IF
End IF
End IF
If Rcount = 0 Then%><tr>
<% End IF%>
<td valign="top" align="left">
<table border="0" cellpadding="0" cellspacing="0" width="300" bgcolor="#EEDDD7" class="t1">
<%if admin Then%>	<tr><td class="udot" colspan="2" align="right">		
			<input onclick="addsub('<%=rsi.Fields("Image")%>')" type="button" value="Add Sub" name="Sub Image">
			<input onclick="image('<%=rsi.Fields("Image")%>')" type="button" value="Image" name="Image">
			<input onclick="edit('<%=rsi.Fields("Image")%>')" type="button" value="Edit" name="Edit"> <input onclick="del('<%=rsi.Fields("Image")%>')" type="button" value="Delete" name="Delete"><input type="checkbox" name="Key" value="<%=rsi.Fields("Image")%>">&nbsp;
			</td></tr><%End IF%>
	<tr>
		<td colspan="2" height="188" class="udot" align="center">
		<a href="Sub_Browser.asp?I=<%=rsi.Fields("Image")%>"><img border="1" src="thumbs/<%=rsi.Fields("Image")%>.jpg"></a></td>
	</tr>
<tr>
				<td colspan="2"><b>Description:</b> 
				<%=rsi.Fields("Descriptions.Name")%></td>
			</tr>
			<tr>
				<td colspan="2"><b>Dimensions: </b><%=rsi.Fields("Width")%> <b>X</b> <%=rsi.Fields("Height")%> 
				<b>X</b> 
				<%=rsi.Fields("Depth")%><font size="2"> in mm</font></td>
			</tr>
			<tr>
				<td colspan="2"><b>Range:</b>&nbsp; <%=rsi.Fields("Ranges.Name")%></td>
			</tr>
			
			<tr>
				<td colspan="2"><b>Colour:</b> 
				<%=rsi.Fields("Colors.Name")%></td>
			</tr>
			<tr>
				<td colspan="2"><b>Edging:</b> <%=rsi.Fields("Edge")%></td>
			</tr>
			<tr>
				<td width="162"><b>Code:</b>&nbsp;&nbsp;&nbsp; <%=rsi.Fields("Code")%>&nbsp; </td>
				<td height="22" width="134"><b>Price:</b> <%=rsi.Fields("Price")%></td>
			</tr>
			<tr>
				<td width="162" colspan="2"><b>Client Ref:</b>&nbsp;&nbsp;&nbsp; <%=rsi.Fields("Client")%>&nbsp; </td>
			</tr>

			
			<tr>
				<td  class="udot" height="30"><b><a href="javascript:d('<%=rsi.Fields("Image")%>');">Info - click 
				here</a></b>
				
				</td>
				<td  class="udot" height="30" align="right">
				
				
				
<%if email Then%>			
			<div style="position: relative; width: 54px; height: 26px; z-index: 1; left: 2px; top: -322px" id="layer1">
			<input onclick="parent.AddImage('<%=rsi.Fields("Image")%>','<%=rsi.Fields("Code")%>');" type="button" value="Email" name="Add"></div><%End IF%>
			</td>
			</tr>
			<tr>
				<td colspan="2">
				<table border="0" cellpadding="0" cellspacing="0" width="295" style="display:none" id="i<%=rsi.Fields("Image")%>">
					<tr>
						<td><%=rsi.Fields("Info")%></td>
					</tr>
				</table>
				</td>
			</tr>
		
</table>
</td>
<%
RCount = RCount + 1
If RCount >= Rows Then
RCount = 0%>
</tr>
<%End IF%>

<%
rsi.MoveNext
Loop
rsi.Close()
%>
<%If RCount < Rows Then%>
<td colspan="<%=(Rows-Rcount)%>"></td>
</tr>
<%End IF%>
</table>
</td></tr>
</table>
</div>
<%if admin Then%>
</form>
<%End IF%>
</td></tr>
</table>
<%
conn.Close ' Close Connection
Set conn = Nothing

%>