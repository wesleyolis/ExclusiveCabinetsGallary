<%
Function List_Ranges(sel)

Set rs = conn.Execute("SELECT Ranges.Range, Ranges.Name FROM Ranges ORDER BY Ranges.Name;")
%>
<select class="input" id="x_Range" name="x_Range" size="1" tabindex="3">
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