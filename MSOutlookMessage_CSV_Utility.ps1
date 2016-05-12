<#
Create a csv file with the WFMPromotions\CSVFiles columns
Populate the FullPath column based on files from a given directory
#>

#### Begin Setup ####

# Get-Process excel | Select-Object -Property Path
# Remove-Item $csv

$folder = "C:\tmp\budgie\G\APR01\forms"
$csv = "c:\tmp\foo.csv"
$xlapp = "C:\Program Files\Microsoft Office\Office15\EXCEL.EXE"

If ( -not $(Test-Path $csv) ) { 
@"
dlist,FullName,BrandID
"@ | Out-File $csv -Encoding "ASCII" }

####  End  Setup ####

#### Begin Main  ####

Get-ChildItem $folder | Select-Object -Property FullName | Export-CSV -Path $csv -Append -Force

Start-Process $xlapp $csv

####  End  Main  ####