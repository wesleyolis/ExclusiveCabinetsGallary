<%
Function List_Colours(sel)

Set rs = conn.Execute("SELECT Colors.Color, Colors.Name FROM Colors ORDER BY Colors.Name;")
%>
<select class="input" id="x_Color" name="x_Color" size="1" tabindex="4">
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