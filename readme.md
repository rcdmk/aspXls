#aspXLS 1.0

A classic ASP Excel writer that currently supports building and writing CSV, TSV (tab separated values)
and HTML outputing, including a pretty print for HTML.

There are plans of writing to XLS and XLSX formats in a near future and reading capabilities for each
of the supported file formats: CSV, TSV, XLS and XLSX.

##License

**The MIT License (MIT) - http://opensource.org/licenses/MIT**
Copyright (c) 2012 RCDMK <rcdmk@rcdmk.com>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT
OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


##Easy to use

Instantiate the class:

    set xls = new aspExl

Add the header/titles for your structure:
	
    ' Add a header: setHeader(x, value)
    xls.setHeader 0, "id"
    xls.setHeader 1, "description"
    xls.setHeader 2, "createdAt"
    
Add some data:

    ' Add the first data row: setValue(x, y, value)
    xls.setValue 0, 0, 1
    xls.setValue 1, 0, "obj 1"
    xls.setValue 2, 0, date()
    
    ' Add a range of values at once: setRange(x, y, valuesArray)
    xls.setRange 0, 2, Array(2, "obj 2", #11/25/2012#)
   
 
**Easy for outputing:**
	
Output the data in string formatted values:
	
    outputCSV = xls.toCSV()
    outputTSV = xls.toTabSeparated()
    
    outputHTML = xls.toHtmlTable()
    
    xls.prettyPrintHTML = true
    outputPrettyHTML = xls.toHtmlTable()
    

Or write it directly to a file:

	' Write the output to a file: writeToFile(filePath, format)
	xls.writeToFile("c:\mydata.csv", ASPXLS_CSV)
	
The format flags supported are:
	
	ASPXLS_CSV = 1	' CSV format
	ASPXLS_TSV = 2	' Tab separeted format
    ASPXLS_HTML = 3	' HTML table format