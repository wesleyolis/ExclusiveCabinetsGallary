<%@ language=vbscript %>
<%
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "no-store"
Response.CacheControl = "no-cache"
%>
<%
Response.Expires = -1
Response.Buffer = true
Response.Clear

Width = 495
Height = 395


Set Image = Server.CreateObject("csImageFile.Manage")
Image.ReadFile Server.MapPath("temp") & "/" & Session("FileName")

'Image.ReadFile Server.MapPath("orig") & "/" & Request.QueryString("Key") & ".jpg"

Image.JpegQuality = 95
Image.FilterType = 5
Image.Resample = True

Image.WriteFile Server.MapPath("orig") & "/" & Request.QueryString("Key") & ".jpg"


'calculate new image size
IF (Width > 0) And (Height > 0) And (Image.Width > Width) And ( Image.Height > Height) Then
	
		If (Image.Width / Image.Height) > (Width / Height) Then
		Image.Resize Width , 0
		Else
		Image.Resize 0 , height
		End IF 
	
Else
IF (Width > 0) And (Image.Width > Width) Then
Image.Resize Width,0
Else
If  (Height > 0) And ( Image.Height > Height) Then
Image.Resize 0,Height
End IF
End If
End IF



'
Image.WriteFile Server.MapPath("temp") & "/thumblarge.jpg"
Call tags
Call graphic
Image.MergeBack Server.MapPath("images") & "\logo.gif",0,0

Image.WriteFile  Server.MapPath("thumbs") & "\" & Request.QueryString("Key") & "_Large.jpg"



'small thumb nail produtrion
Image.ReadFile Server.MapPath("temp") & "/thumblarge.jpg"

'Image.ReadFile Server.MapPath("orig") & "/" & Request.QueryString("Key") & ".jpg"

Image.JpegQuality = 95
Image.FilterType = 5
Image.Resample = True


Width = 300
Height = 180

'calculate new image size
IF (Width > 0) And (Height > 0) And (Image.Width > Width) And ( Image.Height > Height) Then
	
		If (Image.Width / Image.Height) > (Width / Height) Then
		Image.Resize Width , 0
		Else
		Image.Resize 0 , height
		End IF 
	
Else
IF (Width > 0) And (Image.Width > Width) Then
Image.Resize Width,0
Else
If  (Height > 0) And ( Image.Height > Height) Then
Image.Resize 0,Height
End IF
End If
End IF

Call tags
Call graphic
Image.MergeBack Server.MapPath("images") & "\logo.gif",0,0

Image.WriteFile  Server.MapPath("thumbs") & "\" & Request.QueryString("Key") & ".jpg"

Image.Delete Server.MapPath("temp") & "/" & Session("FileName")
Image.Delete Server.MapPath("temp") & "/thumblarge.jpg"
Sub tags

'write image tags for copyright

FFO_CopyrightFlag = True
FFO_Marked = True

FFO_Title = "Exclusive Cabinets Office Furniture ©"
FFO_Author = "Exclusive Cabinets ©  Tel:27-21-9053770"
FFO_Caption = "Office Furniture ©"
FFO_Category = "Office Furniture"
FFO_CopyrightNotice = "The Rights of Editing or Changing in part or hole of this file is prohibated. Missleading this file to be or appear as one's own content is also Prohibated. Any infligment of this copyright, shall lead to legal action againts the parties involved"
FFO_ImageURL = "www.exclusivecabinets.co.za"

FFO_KeywordsAdd = "Exclusive"
FFO_KeywordsAdd = "Cabinets"
FFO_KeywordsAdd = "Office"
FFO_KeywordsAdd = "Furniture"
FFO_KeywordsAdd = "Cupboards"


End Sub

Sub graphic

'import copright infomation
image.Antialias = True
image.TextOpaque = false
image.TextColor = "FFFFFF"
Image.TextSize = 14
Image.TextBold = true
Image.Text image.width - 230, image.height -35, "www.ExclusiveCabinets.co.za"
Image.Text image.width - 150, image.height -20, "Tel: 27-21-9053770"
Image.TextSize = 16
Image.Text image.width-20, image.height -35,  "©"
'impose logo over image
Image.Tile = true

Image.Transparent = True
Image.TransPercent = 98


End Sub

%>