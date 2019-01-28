<%

' Advanced User Level Security for ASPMaker 5+
Const ewAllowAdd = 1
Const ewAllowDelete = 2
Const ewAllowEdit = 4
Const ewAllowView = 8
Const ewAllowList = 8
Const ewAllowReport = 8
Const ewAllowSearch = 8
Const ewAllowAdmin = 16
Dim arUserLevel ' User Level definitions
Dim arUserLevelPriv ' User Level privileges

' Define User Level Variables
Dim ewCurLvl ' Current user level
ewCurLvl = CurrentUserLevel()
Dim ewCurSec

' No user level security
Sub SetUpUserLevel
End Sub

' Get current user privilege
Function CurrentUserLevelPriv(TableName)
	CurrentUserLevelPriv = GetUserLevelPrivEx(TableName, CurrentUserLevel)
End Function

' Get anonymous user privilege
Function GetAnonymousPriv(TableName)
	GetAnonymousPriv = GetUserLevelPrivEx(TableName, 0)
End Function

' Get user privilege based on table name and user level
Function GetUserLevelPrivEx(TableName, UserLevel)
	GetUserLevelPrivEx = 0
	If CStr(UserLevel) = "-1" Then ' System Administrator
		GetUserLevelPrivEx = 31
	ElseIf UserLevel >= 0 Then
		If IsArray(arUserLevelPriv) Then
			Dim I
			For I = 0 to UBound(arUserLevelPriv, 2)
				If CStr(arUserLevelPriv(0, I)) = CStr(TableName) And _
					CStr(arUserLevelPriv(1, I)) = CStr(UserLevel) Then
					GetUserLevelPrivEx = arUserLevelPriv(2, I)
					If IsNull(GetUserLevelPrivEx) Then GetUserLevelPrivEx = 0
					If Not IsNumeric(GetUserLevelPrivEx) Then GetUserLevelPrivEx = 0
					GetUserLevelPrivEx = CLng(GetUserLevelPrivEx)
					Exit For
				End If
			Next
		End If
	End If
End Function

' Get current user level name
Function CurrentUserLevelName
	GetUserLevelName(CurrentUserLevel)
End Function

' Get user level name based on user level
Function GetUserLevelName(UserLevel)
	GetUserLevelName = ""
	If CStr(UserLevel) = "-1" Then
		GetUserLevelName = "Administrator"
	ElseIf UserLevel >= 0 Then
		If IsArray(arUserLevel) Then
			Dim I
			For I = 0 to UBound(arUserLevel, 2)
				If CStr(arUserLevel(0, I)) = CStr(UserLevel) Then
					GetUserLevelName = arUserLevel(1, I)
					Exit For
				End If
			Next
		End If
	End If
End Function

' Sub to display all the User Level settings (for debug only)
Sub ShowUserLevelInfo
	Dim I
	If IsArray(arUserLevel) Then
		Response.Write "User Levels:<br>"
		Response.Write "UserLevelId, UserLevelName<br>"
		For I = 0 To UBound(arUserLevel, 2)
			Response.Write "&nbsp;&nbsp;" & arUserLevel(0, I) & ", " & _
				arUserLevel(1, I) & "<br>"
		Next
	Else
		Response.Write "No User Level definitions." & "<br>"
	End If
	If IsArray(arUserLevelPriv) Then
		Response.Write "User Level Privs:<br>"
		Response.Write "TableName, UserLevelId, UserLevelPriv<br>"
		For I = 0 To UBound(arUserLevelPriv, 2)
			Response.Write "&nbsp;&nbsp;" & arUserLevelPriv(0, I) & ", " & _
				arUserLevelPriv(1, I) & ", " & arUserLevelPriv(2, I) & "<br>"
		Next
	Else
		Response.Write "No User Level privilege settings." & "<br>"
	End If
	Response.Write "CurrentUserLevel = " & CurrentUserLevel & "<br>"
End Sub

' Function to check privilege for List page (for menu items)
Function AllowList(TableName)
	AllowList = CBool(CurrentUserLevelPriv(TableName) And ewAllowList)
End Function

' Get current user name from session
Function CurrentUserName
	CurrentUserName = Session(ewSessionUserName) & ""
End Function

' Get current user id from session
Function CurrentUserID
	CurrentUserID = Session(ewSessionUserID) & ""
End Function

' Get current parent user id from session
Function CurrentParentUserID
	CurrentParentUserID = Session(ewSessionParentUserID) & ""
End Function

' Get current user level from session
Function CurrentUserLevel
	If IsLoggedIn Then
		CurrentUserLevel = Session(ewSessionUserLevel)
	Else
		CurrentUserLevel = 0 ' Anonymous if not logged in
	End If
End Function

' Check if user is logged in
Function IsLoggedIn
	IsLoggedIn = (Session(ewSessionStatus) = "login")
End Function

' Check if user is system administrator
Function IsSysAdmin
	IsSysAdmin = (Session(ewSessionSysAdmin) = 1)
End Function

' Save user level to session
Sub SaveUserLevel
	Session(ewSessionArUserLevel) = arUserLevel
	Session(ewSessionArUserLevelPriv) = arUserLevelPriv
End Sub

' Load user level from session
Sub LoadUserLevel
	If Not IsArray(Session(ewSessionArUserLevel)) Then
		SetupUserLevel
		SaveUserLevel
	End If
	arUserLevel = Session(ewSessionArUserLevel)
	arUserLevelPriv = Session(ewSessionArUserLevelPriv)
End Sub
%>
