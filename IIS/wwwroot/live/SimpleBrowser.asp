<%@ LANGUAGE = VBScript %>
<!--METADATA TYPE="typelib" 
uuid="00000206-0000-0010-8000-00AA006D2EA4" -->
<% Session.Timeout = 20 %>
<%
Response.ContentType = "text/html; charset=windows-1252"
Response.expires = 43200
Response.expiresabsolute = Now() + 100
Response.addHeader "pragma", "cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "cache"
%>
<!--#include file="db.asp"-->

<%
Response.Buffer = true
Dim dir,udir,str,email,admin, rs,rsi,FilRange, FIL,Group1,Group2, Range,Description, Page,Rows,RCount,Rangetxt,Destxt

Rangetxt =""
Destxt =""
Page = 1
Group1 =""
Group2 =""
Range = ""
Description = ""
FIL=""
FilRange = -1
str=""
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

if Request.QueryString("Range").Count > 0 Then
Range = Request.QueryString("Range")
IF (Not Range = "-1") AND Range >= 0 And Range < 65536 Then
FIL = FIL &  " AND (([Images.Range])=" & Range & ") "

sSql = "SELECT * FROM [Ranges] Where [Range] =" & Range
Set rs = conn.Execute(sSql)
if Not rs.eof Then
Rangetxt = rs("Name")
End IF
rs.Close()

End IF
End IF

if Request.QueryString("Des").Count > 0 Then
Description = Request.QueryString("Des")
IF  (Not Description = "-1") AND  Description >= 0 And Description < 65536 Then
FIL = FIL &  " AND (([Images.Description])=" & Description & ") "
sSql = "SELECT * FROM [Descriptions] Where [Description] =" & Description
Set rs = conn.Execute(sSql)
if Not rs.eof Then
Destxt = rs("Name")
End IF
rs.Close()

End IF
End IF

if Request.QueryString("Color").Count > 0 Then
Color = Request.QueryString("Color")
IF  (Not Color = "-1") AND  Color >= 0 And Color < 65536 Then
FIL = FIL &  " AND (([Images.Color])=" & Color & ") "
END IF
End IF

if Request.QueryString("d").Count = 0 Then
dir = -2
else

if (Request.QueryString("d")>= -1) And (Request.QueryString("d") < 655536) Then
dir = Request.QueryString("d")

End IF
If Not dir = "-2" Then
FIL = FIL & "AND(([Images.Dir]= "&dir&"))"



END IF
End IF

if Request.QueryString("Ulti")<> "" Then
Ulti = Request.QueryString("Ulti")
FIL = FIL & " AND (InStr(UCase(([Images.Code])), UCase('" & Ulti & "')) > 0)"       '  " AND (UCase(([Images.Code]))=UCase('" & Ulti & "')) "    (InStr(UCase(([Images.Code])), UCase('" & Ulti & "')) > 0)
End IF

'Response.write(FIL)
sql = "SELECT Images.Image, Images.Code, Images.Client, Ranges.Range, Ranges.Name, Descriptions.Description ,Descriptions.Name, Colors.Name, Images.Width, Images.Height, Images.Depth,IIF([Images.Price]=0,'P.O.A','R ' & [Images.Price]) As Price,Images.Edge, Images.Info"_
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
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>
<link rel="stylesheet" href="gallery.css" type="text/css">
<style>
 .udot {border-bottom-style: dashed; border-bottom-width: 1px; padding: 2px; }
 .t1{font-size:11pt;border: 2px solid #cc907e; padding: 0}
 .GHeading{border-bottom: 2px solid #FF0000; font-size:14pt; color:#DE4010; font-weight:bold}
</style>

<body background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed">

<script type="text/javascript">
<!--
function d(img)
{
if(document.getElementById("i"+ img).style.display == 'block')
document.getElementById("i"+ img).style.display = 'none';
else
document.getElementById("i"+ img).style.display = 'block';
}
//-->
</script>


	<div align="center">
	<table width="800px">
<%
if Not (Request.QueryString("Des").Count > 0 And Request.QueryString("Range").Count > 0)  Then
if Request.QueryString("Des").Count > 0 Then 
%>
<tr><td align="center"><font color="#DE4010" size="5"><b><%=Destxt%></b></font></td></tr>
<%
Else%>
<tr><td align="center"><font color="#DE4010" size="5"><b><%=Rangetxt%></b></font></td></tr>
<%
End IF
Else%>
<tr><td align="center"><font color="#DE4010" size="5"><b><%=Rangetxt%>&nbsp;-&nbsp;<%=Destxt%></b></font></td></tr>
<%End IF%>
<tr><td align="center" colspan="2">

<%

If rsi.eof Then
%><a href="../../welcome.htm">Home</a>&nbsp;&nbsp;| &nbsp;<a href="../welcome.htm">Gallery</a><br><font color="#DE4010" size="4"><b>No Results Found</b></font>
<%
Else%>
<tr height="10px"><td></td></tr>
<tr><td align="center"><a href="../../welcome.htm">Home</a> &nbsp;|&nbsp;&nbsp;<a href="../welcome.htm">Gallery</a><br>
<%
if Not (Request.QueryString("Des").Count > 0 And Request.QueryString("Range").Count > 0)  Then
if Request.QueryString("Des").Count > 0  Then
Set rs = conn.Execute("SELECT Ranges.Range, Ranges.Name FROM Images INNER JOIN Ranges ON Images.Range = Ranges.Range GROUP BY Ranges.Range, Ranges.Name, Images.Description HAVING (((Images.Description)=" & Description & ")) ORDER BY Ranges.Name;")

Do While Not rs.Eof
%>
<a href="#r<%=rs("Range")%>"><%=rs("Name")%></a> &nbsp;&nbsp;&nbsp;&nbsp;
<%
rs.MoveNext
Loop
rs.close()
Else
IF Request.QueryString("Range").Count > 0  Then
Set rs = conn.Execute("SELECT Descriptions.Description, Descriptions.Name FROM Descriptions INNER JOIN Images ON Descriptions.Description = Images.Description GROUP BY Descriptions.Description, Descriptions.Name, Images.Range HAVING (((Images.Range)=" & Range & ")) ORDER BY Descriptions.Name;")

Do While Not rs.Eof
%>
<a href="#d<%=rs("Description")%>"><%=rs("Name")%></a> &nbsp;&nbsp;&nbsp;&nbsp;
<%
rs.MoveNext
Loop
rs.close()

End IF
End IF
End IF

%>
</td></tr>
<%
End IF
Rows = 2
RCount = 0
%>
</p>
<table border="0" cellpadding="0" cellspacing="20" width="800px">
<%
Do While Not rsi.eof 'And (j < rsi.PageSize )
'j = j + 1
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
RCount=0
%><td colspan="<%=(Rows-Rcount)%>">&nbsp;</td>
<% End IF

If (Range = "" And Description ="") Then
%>
<tr><td colspan="<%=Rows%>" class="Gheading"><%=Group1%>&nbsp;-&nbsp;<%=Group2%>&nbsp;</td></tr>
<%
Else
If (Range = "") Then
%><tr><td class="Gheading" colspan="<%=Rows%>"><a name="r<%=rsi.Fields("Range")%>" href=""></a><%=Group1%>&nbsp;</td></tr>
<%
Else%>
<tr><td colspan="<%=Rows%>" class="Gheading"><a name="d<%=rsi.Fields("Description")%>" href=""></a><%=Group2%>&nbsp;</td></tr>
<%
End IF
End IF
End IF

If RCount = 0 Then%><tr>
<% End IF%>
<td valign="top" align="left">
<table border="0" cellpadding="0" cellspacing="0" width="300" bgcolor="#EEDDD7" class="t1">
<tr><td colspan="2" height="188" class="udot" align="center">
<a href="images/html/image.asp?I=<%=rsi.Fields("Image")%>"><img border="1" src="images/thumbs/image.asp?I=<%=rsi.Fields("Image")%>"></a></td>
</tr>
<tr><td colspan="2"><b>Description:</b> <%=rsi.Fields("Descriptions.Name")%></td></tr>
<tr><td colspan="2"><b>Dimensions: </b><%=rsi.Fields("Width")%><b>&nbsp;X&nbsp;</b><%=rsi.Fields("Height")%><b>&nbsp;X&nbsp;</b><%=rsi.Fields("Depth")%><font size="2">&nbsp;mm</font></td></tr>
<tr><td colspan="2"><b>Range:</b>&nbsp;<%=rsi.Fields("Ranges.Name")%></td></tr>
<tr><td colspan="2"><b>Colour:</b>&nbsp;<%=rsi.Fields("Colors.Name")%></td></tr>
<tr><td width="162"><b>Code:</b>&nbsp;&nbsp;&nbsp; <%=rsi.Fields("Code")%>&nbsp;</td><td height="22" width="134"><b>Price:</b> <%=rsi.Fields("Price")%></td></tr>
<tr><td width="162" colspan="2"><b>Gallery Code:</b>&nbsp;&nbsp;&nbsp; <%=rsi.Fields("Image")%>&nbsp; </td></tr>
<tr><td colspan="2"><table border="0" cellpadding="0" cellspacing="0" width="295" style="display:none" id="i<%=rsi.Fields("Image")%>"><tr>
<td><%=rsi.Fields("Info")%></td>
</tr></table></td>
</tr></table>
</td>
<%
RCount = RCount + 1
%><%
If RCount  >= Rows Then
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
<tr><td align="center" colspan="<%=Rows%>"><a href="../index.html">Gallery Home</a><br></td></tr>
</table>
</td></tr>
</table>
</div>
</td></tr>
</table>
	</div>

<%
conn.Close ' Close Connection
Set conn = Nothing

%>