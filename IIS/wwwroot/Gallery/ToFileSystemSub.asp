<%
Response.expires = -1

'If Request.QueryString("I").Count And (Request.QueryString("d") < 65536) And Request.QueryString("Key") > 0 Then
%>
<!--#INCLUDE FILE="clsUpload.asp"-->
<%

Dim objUpload
Dim strFileName, tmpFileName
Dim strPath
Dim Key

Set objUpload = New clsUpload
strPath = Server.MapPath("orig") & "\" & Request.QueryString("I") & "_" & Request.QueryString("Key") & ".jpg"
strFileName = objUpload.Fields("File1").FileName
Session("FileName") = strFileName 
tmpFileName = Server.MapPath("temp") & "\" & strFileName
objUpload("File1").SaveAs tmpFileName
Set objUpload = Nothing

Server.Execute("create_thumb_Sub.asp")

Response.Redirect "Sub_Browser.asp?I=" & Request.QueryString("I")
%>
<%
'Else
'Response.Redirect("Error.asp")
'End If
%>
