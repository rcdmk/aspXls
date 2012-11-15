<%
' Classic ASP CSV creator
' By RCDMK <rcdmk@rcdmk.com>
'
' The MIT License (MIT) - http://opensource.org/licenses/MIT
' Copyright (c) 2012 RCDMK <rcdmk@rcdmk.com>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
' associated documentation files (the "Software"), to deal in the Software without restriction,
' including without limitation the rights to use, copy, modify, merge, publish, distribute,
' sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial
' portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
' NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
' NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
' DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT
' OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


class aspExl
	dim lines(), curBoundX, curBoundY
	dim headers()
	
	sub class_initialize()
		curBoundX = -1
		curBoundY = -1
	end sub
	
	sub class_terminate()
		redim lines(-1)
	end sub
	
	private sub resizeCols(byval newSize)
		dim i, cols
		
		for i = 0 to curBoundY
			cols = lines(i)
			redim preserve cols(newSize)
			lines(i) = cols
		next
		
		redim preserve headers(newSize)
		
		curBoundX = newSize
	end sub
	
	private sub resizeRows(byval newSize)
		dim i
		redim preserve lines(newSize)
		
		for i = curBoundY + 1 to newSize
			if i >= 0 then lines(i) = Array()
		next
		
		curBoundY = newSize
		
		resizeCols curBoundX
	end sub
	
	public sub setHeader(byval x, byval value)
		if x > curBoundX then resizeCols(x)
		
		headers(x) = value
	end sub
	
	public sub addValue(byval x, byval y, byval value)
		dim cols
		
		if y > curBoundY then resizeRows y
		
		if x > curBoundX then resizeCols x
		
		cols = lines(y)
		cols(x) = value
		
		lines(y) = cols
	end sub
	
	public sub addRange(byval x, byval y, byval arr)
		if y > curBoundY then resizeRows y
		
		dim arrBound
		arrBound = ubound(arr)
		
		if arrBound + x > curBoundX then resizeCols(arrBound + x)
		
		dim i, cols
		cols = lines(y)
		
		for i = 0 to arrBound
			cols(x + i) = arr(i)
		next
		
		lines(y) = cols
	end sub
	
	private function toString(byval separator)
		dim output, headersString, i
		output = ""
		headersString = join(headers, separator)
		
		if replace(headersString, separator, "") <> "" then output = headersString & vbCrLf
		
		for i = 0 to curBoundY
			output = output & join(lines(i), separator) & vbCrLf
		next
		
		toString = output
	end function
	
	public function toCSV()
		toCSV = toString(";")
	end function
	
	public function toTabSeparated()
		toTabSeparated = toString(vbTab)
	end function
	
	public function toHtmlTable()
		dim output, headersString, i
		output = "<table border=""1"">" & vbCrLf
		headersString = join(headers, "</th><th>")
		
		if replace(headersString, "</th><th>", "") <> "" then output = output & "<tr><th>" & headersString & "</th></tr>" & vbCrLf
		
		for i = 0 to curBoundY
			output = output & "<tr><td>" & join(lines(i), "</td><td>") & "</td></tr>" & vbCrLf
		next
		
		output = output & "<table>"
		
		toHtmlTable = output
	end function
	
	public sub writeToFile(byval filePath)
		dim fso, file
		dim i
		set fso = createObject("scripting.filesyStemObject")
		
		set file = fso.openTextFile(filePath)
		
		file.write toString()
		
		file.close
		set file = nothing
		
		set fso = nothing
	end sub
end class
%>
