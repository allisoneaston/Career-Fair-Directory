from datetime import datetime
import csv
import os
inventory_path="inventory.csv"
today=datetime.today()
selling_close=datetime(today.year,1,10,17,30 )
testing_end=datetime(today.year,1,9,10)
if(today>selling_close):
	print "Price View Authorized"
gen_html_path="."+os.sep+"webcontent"+os.sep+"BookswapInventory.html"
gen_html=open(gen_html_path,'w')
gen_html.write(r'''<!doctype html>
<html lang="en">

<!--Meta information-->
<head>
<meta charset=utf-8>
<title>Tau Beta Pi Book Swap</title>
<link rel="stylesheet" href="Ultramarine.css" type="text/css">
<link rel="stylesheet" href="bookswap.css" type="text/css">
<link rel="stylesheet" href="inventory.css" type="text/css">
<meta name=description value="Tau Beta Pi hosts a book swap during the first week of class each semester. Students can buy and sell books for their engineering and prerequisite courses">
<meta name=keywords value="book, textbook, tbp, bookswap, engineering, book swap">

<!--[if lt IE 9]>
<script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
<![endif]-->

<script type="text/javascript">
var pagename="Inventory";
</script>
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-26073977-3']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
</head>

<!--Main Content-->
<body>

<!--#include virtual="includes/header.html" -->
<article>
<h1>Current Inventory</h1>''')
if(today<=selling_close):
	gen_html.write(r'''<p>Please note that until reception of books is closed, no price information will be visible.</p>''')
if(today<testing_end):
	gen_html.write(r'''<p><b>This webpage is still under development. Inventory records will not be real data until Book Swap opens. </b></p>''')
gen_html.write(r'''<p>List last updated on: '''+today.strftime("%A, %d %b. at %I:%M%p")+r'''.</p>''')
gen_html.write(r'''<p>Click on column headers to sort the table by that column. Type in boxes to filter.</p>
<table class="table-autosort:0">''')
if(today>selling_close):
	gen_html.write(r'''<thead><tr>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by course name">Course</th>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by title">Book Title</th>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by author">Author</th>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by edition">Edition</th>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by ISBN">ISBN</th>
<th class="table-sortable:numeric table-filterable" title="Click to sort by Quantity">Quantity</th>
<th class="table-sortable:currency table-filterable" title="Click to sort by minimum average">Minimum Price</th>
<th class="table-sortable:currency table-filterable" title="Click to sort by average price">Average Price</th>
</tr>
<tr>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><select onchange="Table.filter(this,this)">
	<option value="function(){return true;}">All</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<10;}"><$10</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<20;}"><$20</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<30;}"><$30</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<40;}"><$40</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<50;}"><$50</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<100;}"><$100</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))>=100;}">>=100</option>
</th>
<th><select onchange="Table.filter(this,this)">
	<option value="function(){return true;}">All</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<10;}"><$10</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<20;}"><$20</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<30;}"><$30</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<40;}"><$40</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<50;}"><$50</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))<100;}"><$100</option>
	<option value="function(val){return parseFloat(val.replace(/\$/,''))>=100;}">>=100</option>
</th>
</tr>
</thead><tbody>''')
else:
	gen_html.write(r'''<thead><tr>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by course name">Course</th>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by title">Book Title</th>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by author">Author</th>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by edition">Edition</th>
<th class="table-sortable:alphanumeric table-filterable" title="Click to sort by ISBN">ISBN</th>
<th class="table-sortable:numeric table-filterable" title="Click to sort by quantity">Quantity</th>
</tr>
<tr>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
<th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
</tr>
</thead><tbody>''')
try:
	inventory_reader = csv.reader(open(inventory_path,'r'), delimiter=',')
	inventory_reader.next()
	for row in inventory_reader:
		if(len(row)<8):
			continue
		gen_html.write(r'''<tr><td>%(Course)s</td>
<td>%(Title)s</td>
<td>%(Author)s</td>
<td>%(Edition)s</td>
<td>%(ISBN)s</td>
<td>%(Quantity)s</td>'''%{"Course":row[0].upper(),"Title":row[1],"Author":row[2],"Edition":row[3],"ISBN":row[4],"Quantity":row[5]})
		if(today>selling_close):
			gen_html.write(r'''<td>%s</td><td>%s</td>'''%(row[6],row[7]))
			gen_html.write("</tr>")
except csv.Error as e:
	print "Formatting Error"
except StopIteration as s:
	print "Empty Inventory"


gen_html.write(r'''</tbody>
</table>


</article>
<!--#include virtual="includes/sidebar.html" -->
<script src="sort_filter_table.js"></script>
</body>
</html>''')
