<% 
' File upload functions for ASPMaker 5+
' (C) 2006 e.World Technology Ltd.

' Config for file upload
Const EW_UploadDestPath = "" ' upload destination path
Const EW_UploadAllowedFileExt = "gif,jpg,jpeg,bmp,png,doc,xls,pdf,zip" ' allowed file extensions

' Function to return path of the uploaded file
'	Parameter: If PhyPath is true(1), return physical path on the server;
'	           If PhyPath is false(0), return relative URL
Function ewUploadPathEx(PhyPath, DestPath)
	Dim Pos
	If PhyPath Then
		ewUploadPathEx = Request.ServerVariables("APPL_PHYSICAL_PATH")
		ewUploadPathEx = ewIncludeTrailingDelimiter(ewUploadPathEx, PhyPath)
		ewUploadPathEx = ewUploadPathEx & Replace(DestPath, "/", "\")
	Else
		ewUploadPathEx = Request.ServerVariables("APPL_MD_PATH")
		Pos = InStr(1, ewUploadPathEx, "Root", 1)
		If Pos > 0 Then	ewUploadPathEx = Mid(ewUploadPathEx, Pos+4)
		ewUploadPathEx = ewIncludeTrailingDelimiter(ewUploadPathEx, PhyPath)
		ewUploadPathEx = ewUploadPathEx & DestPath
	End If
	ewUploadPathEx = ewIncludeTrailingDelimiter(ewUploadPathEx, PhyPath)
End Function

' Function to change the file name of the uploaded file
Function ewUploadFileNameEx(Folder, FileName)
	Dim OutFileName
	
	' By default, ewUniqueFileName() is used to get an unique file name.
	' Amend your logic here
	OutFileName = ewUniqueFileName(Folder, FileName)

	' Return computed output file name
	ewUploadFileNameEx = OutFileName
End Function

' Function to return path of the uploaded file
' returns global upload folder, for backward compatibility only
Function ewUploadPath(PhyPath)
	ewUploadPath = ewUploadPathEx(PhyPath, EW_UploadDestPath)
End Function

' Function to change the file name of the uploaded file
' use global upload folder, for backward compatibility only
Function ewUploadFileName(FileName)
	ewUploadFileName = ewUploadFileNameEx(ewUploadPath(True), FileName)
End Function

' Function to generate an unique file name (filename(n).ext)
Function ewUniqueFileName(Folder, FileName)
	If FileName = "" Then FileName = ewDefaultFileName()

	If FileName = "." Then
		Response.Write "Invalid file name: " & FileName
		Response.End
		Exit Function
	End If
	
	If Folder = "" Then
		Response.Write "Unspecified folder"
		Response.End
		Exit Function
	End If
	
	Dim Name, Ext, Pos
	Name = ""
	Ext = ""
	Pos = InStrRev(FileName, ".")
	If Pos = 0 Then
		Name = FileName
		Ext = ""
	Else
		Name = Mid(FileName, 1, Pos-1)
		Ext = Mid(FileName, Pos+1)
	End If
	
	Folder = ewIncludeTrailingDelimiter(Folder, True)
	
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	If Not fso.FolderExists(Folder) Then
		If Not ewCreateFolder(Folder) Then
			Response.Write "Folder does not exist: " & Folder
			Set fso = Nothing
			Exit Function
		End If
	End If
	
	Dim Suffix, Index
	Index = 0
	Suffix = ""
	
	' Check to see if filename exists
	While fso.FileExists(folder & Name & Suffix & "." & Ext)
		Index = Index + 1
		Suffix = "(" & Index & ")"
	Wend
	Set fso = Nothing

	' Return unique file name
	ewUniqueFileName = Name & Suffix & "." & Ext
	
End Function

' Function to create a default file name (yyyymmddhhmmss.bin)
Function ewDefaultFileName
	Dim DT
	DT = Now()
	ewDefaultFileName = ewZeroPad(Year(DT), 4) & ewZeroPad(Month(DT), 2) &  _
		ewZeroPad(Day(DT), 2) & ewZeroPad(Hour(DT), 2) & _
		ewZeroPad(Minute(DT), 2) & ewZeroPad(Second(DT), 2) & ".bin"
End Function

' Function to check the file type of the uploaded file
Function ewUploadAllowedFileExt(FileName)
	If Trim(FileName & "") = "" Then
		ewUploadAllowedFileExt = True
		Exit Function
	End If
	Dim Ext, Pos, arExt, FileExt
	arExt = Split(EW_UploadAllowedFileExt & "", ",")
	Ext = ""
	Pos = InStrRev(FileName, ".")
	If Pos > 0 Then	Ext = Mid(FileName, Pos+1)
	ewUploadAllowedFileExt = False
	For Each FileExt in arExt
	  If LCase(Trim(FileExt)) = LCase(Ext) Then
	    ewUploadAllowedFileExt = True
	    Exit For
	  End If
	Next
End Function

' Function to include the last delimiter for a path
Function ewIncludeTrailingDelimiter(Path, PhyPath)
	If PhyPath Then
		If Right(Path, 1) <> "\" Then Path = Path & "\"
	Else
		If Right(Path, 1) <> "/" Then Path = Path & "/"
	End If
	ewIncludeTrailingDelimiter = Path
End Function

' Function to write the paths for config/debug only
Sub ewWriteUploadPaths
	Response.Write "Request.ServerVariables(""APPL_PHYSICAL_PATH"")=" & _
		Request.ServerVariables("APPL_PHYSICAL_PATH") & "<br>"
	Response.Write "Request.ServerVariables(""APPL_MD_PATH"")=" & _
		Request.ServerVariables("APPL_MD_PATH") & "<br>"
End Sub 

'===============================================================================
' Other functions for file upload (Do not modify)

Function stringToByte(toConv)
	Dim i, tempChar
	For i = 1 to Len(toConv)
		tempChar = Mid(toConv, i, 1)
		stringToByte = stringToByte & ChrB(AscB(tempChar))
	Next
End Function

Private Function ByteToString(ToConv)
	On Error Resume Next
 	For I = 1 to LenB(ToConv)
 	  ByteToString = ByteToString & Chr(AscB(MidB(ToConv,i,1)))
 	Next
End Function

Function ConvertToBinary(RawData)
	Dim oRs
	Set oRs = Server.CreateObject("ADODB.Recordset")
	' Create field in an empty RecordSet
	Call oRs.Fields.Append("Blob", 205, LenB(RawData)) ' Add field with type adLongVarBinary
	Call oRs.Open()
	Call oRs.AddNew()
	Call oRs.Fields("Blob").AppendChunk(RawData & ChrB(0))
	Call oRs.Update()
	' Save Blob Data
	ConvertToBinary = oRs.Fields("Blob").GetChunk(LenB(RawData))
	' Close RecordSet
	Call oRs.Close()
	Set oRs = Nothing
End Function

Function ConvertToUnicode(RawData)
	Dim oRs		
	Set oRs = Server.CreateObject("ADODB.Recordset")
	' Create field in an empty recordset
	Call oRs.Fields.Append("Text", 201, LenB(RawData)) ' Add field with type adLongVarChar
	Call oRs.Open()
	Call oRs.AddNew()
	Call oRs.Fields("Text").AppendChunk(RawData & ChrB(0))
	Call oRs.Update()
	' Save Unicode Data
	ConvertToUnicode = oRs.Fields("Text").Value
	' Close recordset
	Call oRs.Close()
	Set oRs = Nothing
End Function

Function getValue(dict, name)
	Dim gv
	If dict.Exists(name) Then
		gv = CStr(dict(name).Item("Value"))
		gv = Left(gv, Len(gv)-2)
		getValue = gv
	Else
		getValue = ""
	End If
End Function

Function getFileData(dict, name)
	If dict.Exists(name) Then
		getFileData = dict(name).Item("Value")
		If LenB(getFileData) Mod 2 = 1 Then
			getFileData = getFileData & ChrB(0)
		End If
	Else
		getFileData = ""
	End If
End Function

Function getFileName(dict, name)
	Dim temp, tempPos
	If dict.Exists(name) Then
		temp = dict(name).Item("FileName")
		tempPos = 1 + InStrRev(temp, "\")
		getFileName = Mid(temp, tempPos)
	Else
		getFileName = ""
	End If
End Function

Function getFileSize(dict, name)
	If dict.Exists(name) Then
		getFileSize = LenB(dict(name).Item("Value"))
	Else
		getFileSize = 0
	End If
End Function

Function getFileContentType(dict, name)
	If dict.Exists(name) Then
		getFileContentType = dict(name).Item("ContentType")
	Else
		getFileContentType = ""
	End If
End Function

Function ewFolderExists(Folder)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	ewFolderExists = fso.FolderExists(Folder)
	Set fso = Nothing
End Function

Sub ewDeleteFile(FilePath)
	On Error Resume Next
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If FilePath <> "" And fso.FileExists(FilePath) Then
		fso.DeleteFile(FilePath)
	End If
	Set fso = Nothing
End Sub

Sub ewRenameFile(OldFilePath, NewFilePath)
	On Error Resume Next
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If OldFilePath <> "" And fso.FileExists(OldFilePath) Then
		fso.MoveFile OldFilePath, NewFilePath
	End If
	Set fso = Nothing
End Sub

Function ewCreateFolder(Folder)
	On Error Resume Next
	ewCreateFolder = False
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(Folder) Then
		If ewCreateFolder(fso.GetParentFolderName(Folder)) Then
			fso.CreateFolder(Folder)
			If Err.Number = 0 Then ewCreateFolder = True
		End If
	Else
		ewCreateFolder = True
	End If
	Set fso = Nothing
End Function

Function ewSaveFile(Folder, FileName, FileData)
	On Error Resume Next
	ewSaveFile = False
	If ewCreateFolder(Folder) Then
		Set oStream = Server.CreateObject("ADODB.Stream")
		oStream.Type = 1 ' 1=adTypeBinary
		oStream.Open
		oStream.Write ConvertToBinary(FileData)
		oStream.SaveToFile Folder & FileName, 2 ' 2=adSaveCreateOverwrite
		oStream.Close
		Set oStream = Nothing
		If Err.Number = 0 Then ewSaveFile = True
	End If
End Function
%>
