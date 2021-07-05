# new-ExcelWorkbook.ps1
# Creates a new workbook (with just one sheet), in Excel 2007
# Then we create and save a sample worksheet
# Thomas Lee 

# Create Excel object

$excel = new-object -comobject Excel.Application
 
# Make Excel visible

$excel.visible = $true
 
# Create a new workbook

$workbook = $excel.workbooks.add()
 
# default workbook has three sheets, remove 2

$S2 = $workbook.sheets | where {$_.name -eq "Sheet2"}
$s3 = $workbook.sheets | where {$_.name -eq "Sheet3"}
$s3.delete()
$s2.delete()
 
# Get sheet and update sheet name

$s1 = $workbook.sheets | where {$_.name -eq 'Sheet1'}
$s1.name = "PowerShell Sample"
 
# Update workook properties

$workbook.author = "Thomas Lee - tfl@psp.co.uk"
$workbook.title = "Excel and PowerShell rock!"
$workbook.subject = "Demonstrating the Power of PowerShell"
 
# Next update some cells in the worksheet 'PowerShell Sample'

$s1.range("A1:A1").cells="Cell a1"
$s1.range("A2:A2").cells="A2"
$s1.range("b1:b1").cells="Cell B1"
$s1.range("b2:b2").cells="b2"
$s1.range("D1:D1").cells=2
$s1.range("D2:D2").cells=2
$s1.range("D3:D3").cells.formula = "=sum(d1,d2)"
 
# And save it away:

$s1.saveas("c:\scripts\test.xlsx")
