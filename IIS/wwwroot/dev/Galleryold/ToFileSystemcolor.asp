<%
Response.expires = -1

'If  Request.QueryString("Key") > 0 Then
'<!--#include file="db.asp"-->

%>
<%
'Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
'conn.Open xDb_Conn_Str
'sql = "UPDATE [Images] SET Images.Sync_img = True WHERE ((([Images.Image])="  & Request.QueryString("Key") &  "));"
'conn.Execute(sql)

%>
<!--#INCLUDE FILE="clsUpload.asp"-->
<%

Dim objUpload
Dim strFileName, tmpFileName
Dim strPath
Dim Key

Set objUpload = New clsUpload
strPath = Server.MapPath("orig") & "\Color_" & Request.QueryString("Key") & ".jpg"
strFileName = objUpload.Fields("File1").FileName
Session("FileName") = strFileName 
tmpFileName = Server.MapPath("temp") & "\Color_" & strFileName
objUpload("File1").SaveAs tmpFileName
Set objUpload = Nothing

Server.Execute("create_thumb_color.asp")

Response.Redirect "ColorslistImages.asp"
%>
<%
'Else
'Response.Redirect("Error.asp")
'End If
%>
