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

Sql = "SELECT Sub_Image.* FROM Sub_Image"
sKey = Request.QueryString("Key")

If sKey = "" Or IsNull(sKey) Then
	sKey = Request.Form("Key")(1)
	orig = Request.Form("orig")
	If sKey = "" Or IsNull(sKey) Then
		'Response.Write "Error"	
		Else
		Sql = Sql & " WHERE [Sub_Image].[ID]=" & sKey
		pos = 2
		max =  Request.Form("Key").Count
		Do While (pos <= max)
		Sql = Sql & " OR [Sub_Image].[ID]=" & Request.Form("Key")(pos)
		pos = pos + 1
		Loop
	End IF
	Else
	orig = Request.querystring("o").count
	Sql  = Sql &  " WHERE [Sub_Image].[ID]=" & sKey
End IF

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open xDb_Conn_Str
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open Sql, conn, 1, 2
		
	Do While Not rs.Eof
	img = rs("Image")
	ID = rs("ID")
	DelImageSub img ,ID, orig
	rs.Delete
	
	
	rs.MoveNext
	Loop
	
	rs.Close
	Set rs = Nothing

Response.Redirect "Sub_Browser.asp?I=" & Request.querystring("I")

Function DelImageSub(img,simg,o)
		path = Server.MapPath("thumbs")
		If (objFSO.FileExists(path & "/" & img & "_" & simg & ".jpg"))=true Then
		objFSO.DeleteFile path & "/" & img & "_" & simg & ".jpg",TRUE
		
		End IF
		
		path = Server.MapPath("orig")
		If (objFSO.FileExists(path & "/" & img & "_" & simg & ".jpg"))=true Then
			If orig = 0 Then
			path2 = Server.MapPath("old_orig")
			objFSO.CopyFile path & "/" & img & "_" & simg & ".jpg", path2 & "/" & img & "_" & simg & ".jpg",TRUE
			End IF
			objFSO.DeleteFile path & "/" & img & "_" & simg & ".jpg",TRUE
		End IF
		'Response.write("<br>Image")
End Function


Set objFSO = Nothing
%>
