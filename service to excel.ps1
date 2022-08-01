$a = new-object -comobject excel.application
$a.Visible = $True
$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)
$c.Cells.Item(1,1) = "Service Name"
$c.Cells.Item(1,2) = "Service Status"
$i = 2
get-service | foreach-object{ $c.cells.item($i,1) = $_.name
$c.cells.item($i,2) = $_.status; $i=$i+1}
$b.SaveAs("C:\temp\!temp1\Test.xlsx")
$a.Quit()