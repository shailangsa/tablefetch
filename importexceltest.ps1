Import-Module ImportExcel

$ExcelData = Open-ExcelPackage -path $PSScriptRoot\spreadsheet.xlsx


$worksheet = $ExcelData.Workbook.Worksheets["Sheet1"]  # Access the desired worksheet by name



foreach ($table in $worksheet.Tables) {
    if ($table.Name -eq "Table4") {
        Write-Host "Found Table: $($table.Name)"
		
		Write-Host "Table Range Address: $($table.Range.Address)"
        
 
    }
	
	

# Output the data (if any)
if ($tableData.Count -gt 0) {
    $tableData | Format-Table -AutoSize | Out-Host
} else {
    Write-Host "No data found in the table."
}
	
}
