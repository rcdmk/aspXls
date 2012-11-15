<% option explicit %><!--#include file="../source/aspExl.class.asp" --><!DOCTYPE HTML>
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
          <div class="brand" href="#">aspExl Tests</div>
        </div>
      </div>
    </div>
	<div class="container">
		<div class="row">
			<div class="span12">
				<h1 class="page-header">aspExl Tests</h1>
				<%
				dim startTime, xls
				startTime = timer
				
				set xls = new aspExl
				
				' Add a header
				xls.setHeader 0, "id"
				xls.setHeader 1, "description"
				xls.setHeader 2, "createdAt"
				
				' Add the first data row
				xls.addValue 0, 0, 1
				xls.addValue 1, 0, "obj 1"
				xls.addValue 2, 0, date()
				
				' Add a bigger imcomplete data row
				xls.addValue 3, 1, "Comment"
				
				' Add a range of values at once
				xls.addRange 0, 2, Array(2, "obj 2", #11/25/2012#)
				
				%>
				<h2><code>toCSV()</code></h2>
				<pre><%= xls.toCSV() %></pre>
				
				<h2><code>toHtmlTalbe()</code></h2>
				<%= xls.toHtmlTable() %>
				
				<%	
				set xls = nothing
				%>
			</div>
		</div>
	</div>
	<div class="navbar navbar-inverse navbar-fixed-bottom">
		<div class="navbar-inner">
			<div class="container">
				<p class="text-info pull-right"><i class="icon-time icon-white"></i> Generated in <%= timer - startTime %>s</p>
			</div>
		</div>
	</div>
</body>
</html>