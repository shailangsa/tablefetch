$ExcelData = Open-ExcelPackage -path $PSScriptRoot\spreadsheet.xlsx


$worksheet = $ExcelData.Workbook.Worksheets["Sheet1"]  # Access the desired worksheet by name



foreach ($table in $worksheet.Tables) {
    if ($table.Name -eq "Table4") {
        Write-Host "Found Table: $($table.Name)"
		
		Write-Host "Table Range Address: $($table.Range.Address)"
        
        # Now you can work with the table, e.g., get the range or data
        $tableData = $table.Range.Value
        $tableData
		$tableData | Format-Table -AutoSize | Out-Host
    }
	
	foreach ($row in $table.Range.Rows) {
    # Skip the header row (first row)
    if ($row.RowNumber -gt 1) {
        $rowObj = @{}
        for ($i = 1; $i -le $row.Cells.Count; $i++) {
            $rowObj[$headers[$i - 1]] = $row.Cells[$i].Text
        }
        $tableData += New-Object PSObject -Property $rowObj
    }
	
}# Output the data (if any)
if ($tableData.Count -gt 0) {
    $tableData | Format-Table -AutoSize | Out-Host
} else {
    Write-Host "No data found in the table."
}
	
}
