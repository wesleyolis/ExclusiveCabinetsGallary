<%@ language=vbscript %>
<%
Response.expires = 2880
Response.expiresabsolute = Now() + 2880
Response.addHeader "pragma", "cache"
Response.addHeader "cache-control", "cache"
Response.CacheControl = "cache"
%>
<%



IF Request.QueryString("I").Count > 0 Then
Dim Image
Quality = 90

IF Request.QueryString("Quality").Count > 0 Then
Quality = Request.QueryString("Quality")
End IF

Set Image = Server.CreateObject("csImageFile.Manage")

Image.ReadFile Server.MapPath("orig/" & Request.QueryString("I") & ".jpg")
Image.JpegQuality = Quality

IF Request.QueryString("Resample").Count > 0 Then
Image.FilterType = 5
Image.Resample = True
End IF


IF (Request.QueryString("Width").Count > 0) And (Request.QueryString("Height").Count > 0) Then
	
	If (Image.Width / Image.Height) > (Request.QueryString("Width") / Request.QueryString("Height")) Then
	Image.Resize Request.QueryString("Width") , 0
	Else
	Image.Resize 0 , Request.QueryString("Height")
	End IF 
	
Else
IF Request.QueryString("Width").Count > 0 Then
Image.Resize Request.QueryString("Width"),0
Else
If  Request.QueryString("Height").Count > 0 Then
Image.Resize 0,Request.QueryString("Height")
End IF
End If
End IF

'write image tags for copyright

FFO_CopyrightFlag = True
FFO_Marked = True

FFO_Title = "Exclusive Cabinets Office Furniture �"
FFO_Author = "Exclusive Cabinets �  Tel:27-21-9053770"
FFO_Caption = "Office Furniture �"
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
Image.TextSize = 14
Image.TextBold = true
Image.Text image.width - 230, image.height -35, "www.ExclusiveCabinets.co.za"
Image.Text image.width - 150, image.height -20, "Tel: 27-21-9053770"
Image.TextSize = 16
Image.Text image.width-20, image.height -35,  "�"
'impose logo over image
Image.Tile = true

Image.Transparent = True
Image.TransPercent = 97
Image.MergeBack Server.MapPath("images") & "\logo.gif",0,0


Response.ContentType = "Image/Jpeg"
Response.BinaryWrite Image.JPGData
Response.End
Else
Response.End
End IF
%>