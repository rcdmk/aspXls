<%
option explicit

%><!--#include file="../source/aspExl.class.asp" --><!DOCTYPE HTML>
<html lang="en-US">
<head>
	<meta charset="ISO-8859-1">
	<title>aspExl Tests</title>
	<link href="//netdna.bootstrapcdn.com/twitter-bootstrap/2.2.1/css/bootstrap.min.css" rel="stylesheet">
	<style type="text/css">
		.row-spaced-top {
			margin-top: 20px;
		}
	</style>
</head>
<body>
	<div class="navbar navbar-inverse navbar-static-top">
      <div class="navbar-inner">
        <div class="container">
          <div class="brand">aspExl &diamond; Usage Tests</div>
        </div>
      </div>
    </div>
	<div class="container">
		<div class="row">
			<div class="span12">
				<h1 class="page-header">Examples</h1>
				<%
				dim startTime, xls, i
				startTime = timer
				
				set xls = new aspExl
				
				' Add a header
				xls.setHeader 0, "id"
				xls.setHeader 1, "description"
				xls.setHeader 2, "createdAt"
				
				' Add the first data row
				xls.setValue 0, 0, 1
				xls.setValue 1, 0, "obj 1"
				xls.setValue 2, 0, date()
				
				' Add a bigger imcomplete data row
				xls.setValue 3, 1, "Comment"
				
				' Add a range of values at once
				xls.setRange 0, 2, Array(2, "obj 2", #11/25/2012#)
				
				
				' Add lots of data
				for i = 3 to 1000
					xls.setValue 0, i, i
					xls.setValue 1, i, "obj " & i
					xls.setValue 2, i, now()
				next
				
				
				dim outputCSV, outputTSV, outputHTML, outputPrettyHTML
				
				outputCSV = xls.toCSV()
				outputTSV = xls.toTabSeparated()
				
				outputHTML = xls.toHtmlTable()
				
				xls.prettyPrintHTML = true
				outputPrettyHTML = xls.toHtmlTable()
				
				%>
				<h2><code>toCSV()</code></h2>
				<pre class="pre-scrollable"><%= outputCSV %></pre>
				
				<h2><code>toTabSeparated()</code></h2>
				<pre class="pre-scrollable"><%= outputTSV %></pre>
				
				<h2><code>toHtmlTalbe() - prettyPrintHTML = false (default)</code></h2>
				<pre class="pre-scrollable"><%= Server.HtmlEncode(outputHTML) %></pre>
				
				<h2><code>toHtmlTalbe() - prettyPrintHTML = true</code></h2>
				<pre class="pre-scrollable"><%= Server.HtmlEncode(outputPrettyHTML) %></pre>
				
				<h4>HTML</h4>
				<%= replace(replace(replace(outputHTML, "<table", "<table class=""table table-striped table-bordered"""), "<thead><tr>", "<thead><tr class=""alert-info"">"), "<tbody", "<tbody class=""pre-scrollable""") %>
				
				<%	
				set xls = nothing
				%>
			</div>
		</div>
	</div>
	<div class="navbar navbar-inverse navbar-static-bottom">
		<div class="navbar-inner">
			<div class="container">
				<p></p>
				<p class="text-info pull-right">
					<i class="icon-time icon-white"></i> Generated in <%= clng((timer - startTime) * 1000) %>ms
				</p>
			</div>
		</div>
	</div>
</body>
</html>