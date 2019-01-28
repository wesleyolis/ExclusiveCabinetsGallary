<%@ language=vbscript %>
<%
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "no-store"
Response.CacheControl = "no-cache"
%>
<%
Response.Buffer = true
Response.Clear

'Width = 151
'Height = 117



Set Image = Server.CreateObject("csImageFile.Manage")
Image.ReadFile Server.MapPath("temp") & "/" & Session("FileName")

'Image.ReadFile Server.MapPath("orig") & "/" & Request.QueryString("Key") & ".jpg"

Image.JpegQuality = 95
Image.FilterType = 5
Image.Resample = True


'Width = 1024
'Height = 768

'calculate new image size

'IF (Width > 0) And (Height > 0) And (Image.Width > Width) And ( Image.Height > Height) Then
	
'		If (Image.Width / Image.Height) > (Width / Height) Then
'		Image.Resize Width , 0
'		Else
'		Image.Resize 0 , height
'		End IF 
	
'Else
'IF (Width > 0) And (Image.Width > Width) Then
'Image.Resize Width,0
'Else
'If  (Height > 0) And ( Image.Height > Height) Then
'Image.Resize 0,Height
'End IF
'End If
'End IF


Image.WriteFile Server.MapPath("orig") & "\" & Request.QueryString("I") & "_" & Request.QueryString("Key") & ".jpg"

Image.JpegQuality = 100

Width = 151
Height = 117

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



'import copright infomation
image.Antialias = True
image.TextOpaque = false
image.TextColor = "FFFFFF"
Image.TextSize = 9
Image.TextBold = false
Image.Text 5, image.height -25, "www.ExclusiveCabinets.co.za"
Image.Text 5, image.height -13, "Tel: 27-21-9053770"
Image.TextSize = 10
Image.Text 131, image.height -25,  "©"
'impose logo over image
Image.Tile = true

Image.Transparent = True
Image.TransPercent = 97
Image.MergeBack Server.MapPath("images") & "\logo.gif",0,0

Image.WriteFile  Server.MapPath("thumbs") & "\" & Request.QueryString("I") & "_" & Request.QueryString("Key") & ".jpg"
Image.Delete Server.MapPath("temp") & "/" & Session("FileName")


%>