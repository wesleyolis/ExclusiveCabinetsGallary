<%
Function List_Descriptions(sel)

Set rs = conn.Execute("SELECT Descriptions.Description, Descriptions.Name FROM Descriptions ORDER BY Descriptions.Name;")
%>
<select class="input" id="x_Description" name="x_Description" size="1" tabindex="4">
<%
Do While Not rs.Eof
If rs.Fields(0) = sel Then
%><option selected value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<%Else%>
<option value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<%
End IF
rs.MoveNext
Loop
Response.write("</select>")
End Function
%>