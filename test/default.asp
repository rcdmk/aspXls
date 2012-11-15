<!--#include file="../source/aspExl.class.asp" -->
<%
set xls = new aspExl

' Add a header
xls.addValue 0, 0, "id"
xls.addValue 1, 0, "description"
xls.addValue 2, 0, "createdAt"

' Add the first data row
xls.addValue 0, 1, 1
xls.addValue 1, 1, "obj 1"
xls.addValue 2, 1, date()

' Add a bigger imcomplete data row
xls.addValue 3, 2, "Comment"

' Add a range of values at once
xls.addRange 3, Array(2, "obj 2", #25/11/2012#)

response.write "<pre>" & xls.toString() & "<pre>"

response.write xls.toHtmlTable()

set xls = nothing
%>