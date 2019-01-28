<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>

<!--#include file="db.asp"-->

<%
Response.Buffer = True
Dim Sql,rs
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

Sql = "SELECT * FROM [Images] "
sKey = Request.QueryString("Key")

If sKey = "" Or IsNull(sKey) Then
	sKey = Request.Form("Key")(1)
	orig = Request.Form("orig")
	If sKey = "" Or IsNull(sKey) Then
		'Response.Write "Error"	
		Else
		Sql = Sql & " WHERE [Image]=" & sKey
		pos = 2
		max =  Request.Form("Key").Count
		Do While (pos <= max)
		Sql = Sql & " OR [Image]=" & Request.Form("Key")(pos)
		pos = pos + 1
		Loop
	End IF
	Else
	orig = Request.querystring("o").count
	Sql  = Sql &  " WHERE [Image]=" & sKey
End IF

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open xDb_Conn_Str
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open Sql, conn, 1, 2
		
	Do While Not rs.Eof
	img = rs("Image")
	rs.Delete
	DelImage img , orig
	
	rs.MoveNext
	Loop
	
	rs.Close
	Set rs = Nothing

Response.Redirect "Dirbrowser.asp?d=" & Request.querystring("d")

Function DelImage(img,o)
		path = Server.MapPath("thumbs")
		If (objFSO.FileExists(path & "/" & img & ".jpg"))=true Then
		objFSO.DeleteFile path & "/" & img & ".jpg",TRUE
		
		End IF
		
		path = Server.MapPath("orig")
		If (objFSO.FileExists(path & "/" & img & ".jpg"))=true Then
			If orig = 0 Then
			path2 = Server.MapPath("old_orig")
			objFSO.CopyFile path & "/" & img & ".jpg", path2 & "/" & img & ".jpg",TRUE
			End IF
			objFSO.DeleteFile path & "/" & img & ".jpg",TRUE
		End IF
		Response.write("<br>Image")
End Function
Set objFSO = Nothing
%>
