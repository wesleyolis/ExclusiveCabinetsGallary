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

Set Image = Server.CreateObject("csImageFile.Manage")
Image.ReadFile Server.MapPath("../../Gallery/thumbs/Color_" & Request.QueryString("I") & ".jpg")
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

Response.ContentType = "Image/Jpeg"
Response.BinaryWrite Image.JPGData
Response.End
Else
Response.End
End IF
%>