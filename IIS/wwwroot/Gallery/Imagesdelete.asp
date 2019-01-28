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
Dim Sql,rs, ID,orig

orig = Null
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Sql = ""
Sql_M = "SELECT Images.* FROM Images "
Sql_S = "SELECT  Sub_Image.* FROM Sub_Image "

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
	orig = Request.QueryString("o").count
	Sql  = Sql &  " WHERE [Image]=" & sKey
End IF

	'Delete Sub
	
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open xDb_Conn_Str
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open Sql_S & Sql, conn, 1, 2
		
	Do While Not rs.Eof
		img = rs("Image")
		ID = rs("ID")
		DelImageSub img ,ID, orig
		rs.Delete
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	
	
	
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open xDb_Conn_Str
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open Sql_M & Sql, conn, 1, 2
		
	Do While Not rs.Eof
		img = rs("Image")
		DelImage img , orig
		rs.Delete
		rs.MoveNext
	Loop
	
	rs.Close
	Set rs = Nothing
	
	
	'delete colour ranges

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	sSql = "Select * FROM Image_Groups " & Sql
	rs.Open sSql, conn, 1, 2
	Do While Not rs.Eof
		rs.Delete
		rs.MoveNext
	Loop
	rs.Close


Response.Redirect "Dirbrowser.asp?d=" & Request.querystring("d")

Function DelImage(img,o)
		path = Server.MapPath("thumbs")
		If (objFSO.FileExists(path & "/" & img & ".jpg"))=True Then
		objFSO.DeleteFile path & "/" & img & ".jpg",TRUE
		
		End IF
		If (objFSO.FileExists(path & "/" & img & "_Large.jpg"))=True Then
		objFSO.DeleteFile path & "/" & img & "_Large.jpg",TRUE
		
		End IF

		
		path = Server.MapPath("orig")
		If (objFSO.FileExists(path & "/" & img & ".jpg"))=True Then
			If orig = 0 Then
			path2 = Server.MapPath("old_orig")
			objFSO.CopyFile path & "/" & img & ".jpg", path2 & "/" & img & ".jpg",TRUE
			End IF
			objFSO.DeleteFile path & "/" & img & ".jpg",TRUE
		End IF
		'Response.write("<br>Image")
End Function

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
