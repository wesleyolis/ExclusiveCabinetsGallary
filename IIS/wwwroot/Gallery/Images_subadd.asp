
<%
Response.Expires = 0
Response.ExpiresAbsolute = #1/1/1980# ' Expired
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<!--#include file="db.asp"-->

<%
' decalre varibles

command = Null
image = Null
key = Null
description = Null


' do cheack for forms data

command = Request.Form("command")
If command <> "" Then
	
	'get form data
	
	key = Request.Form("key")
	image = Request.Form("image")
	description = Request.Form("des")
	
	' element exist cheack which type of add, update
	if command = "Add" Then
		If add() Then
		Response.Clear
		Response.Redirect "ImagesSubAdd.asp?I=" & image & "&Key=" & key & "&f="
		End If
	Else
		If edit() Then
		Response.Clear
		Response.Redirect "ImagesSubAdd.asp?I=" & image & "&Key=" & key
		End If
	End if
	
	
Else
' New entry or update

key = Request.QueryString("key")

If key <> "" Then

	'yes then it must be an update
	command="Edit"

	'must now load the data
	If LoadData() Then
	'continue and show form
	End IF
Else
image = Request.QueryString("I")
command="Add"
End IF
End IF

Function LoadData()

	' Open connection to the database

	Sql = "SELECT Sub_Image.* FROM Sub_Image WHERE (ID=" & key & ");"


		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open xDb_Conn_Str
		
		Set rs = conn.Execute(Sql)
		
		If Not rs.eof Then
		'set the elements fields
		
			description = rs("Des")
			image = rs("Image")	
		Else
			'Error the element was not found
			command = "Add"
		End IF

End Function


Function add()
	add = False
	' Open connection to the database
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open xDb_Conn_Str
	
		Sql = "SELECT Sub_Image.* FROM Sub_Image  WHERE (0 = 1);"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3
		rs.Open Sql, conn, 1, 2
		rs.AddNew
		rs("Image") = image
	
		' Field Code
		sTmp = Trim(description)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("Des") = sTmp
		
		rs.Update
		key = rs("ID")
		rs.Close
		Set rs = Nothing
	
	conn.Close
	add = True
	
End Function

Function edit()
	' Open connection to the database
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open xDb_Conn_Str
	Sql = "SELECT * FROM [Sub_Image] WHERE (ID=" & key & ")"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open Sql, conn, 1, 2
	If rs.Eof Then
		edit = False ' Update Failed
	Else
	image = rs("Image")
	' Field Code
	sTmp = Trim(description)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("Des") = sTmp
	
	rs.Update
	rs.Close
	Set rs = Nothing
	conn.Close

	edit = True
	End IF
End Function



%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Add Sub Image</title>
</head>

<body>
<form method="POST" action="Images_subadd.asp">
<table border="0" width="311" id="table1" cellspacing="0" cellpadding="0" height="60">
	<tr>
		<td background="images/bar_back.gif" height="29" width="85"><b>
		Description</b></td>
		<td height="29"><input type="text" name="des" size="30" maxlength="30" value="<%=description%>"></td>
	</tr>
	<tr>
		<td width="85">&nbsp;</td>
		<td align="right"><input type="submit" value="Submit" name="B1"></td>
	</tr>
</table>
<input type="hidden" name="key" value="<%=key%>">
<input type="hidden" name="image" value="<%=image%>">
<input type="hidden" name="command" value="<%=command%>">

	
</form>

</body>

</html>