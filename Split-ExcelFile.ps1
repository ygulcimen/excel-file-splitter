# Excel File Splitter Script
$inputFile = "C:\Users\filePath"  # Change this path
$outputFolder = "C:\Users\FilePath"  # Change this path
$chunkSize = 150000  # Rows per file

# Create output folder if it doesn't exist
if (!(Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder
}

Write-Host "Starting Excel file split process..." -ForegroundColor Green
Write-Host "Input file: $inputFile" -ForegroundColor Yellow
Write-Host "Chunk size: $chunkSize rows per file" -ForegroundColor Yellow

try {
    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "Opening Excel file..." -ForegroundColor Yellow
    $workbook = $excel.Workbooks.Open($inputFile)
    $worksheet = $workbook.Sheets.Item(1)
    
    # Get total rows
    $lastRow = $worksheet.UsedRange.Rows.Count
    $lastCol = $worksheet.UsedRange.Columns.Count
    
    Write-Host "Total rows found: $lastRow" -ForegroundColor Cyan
    Write-Host "Total columns found: $lastCol" -ForegroundColor Cyan
    
    # Calculate number of files needed
    $totalFiles = [math]::Ceiling($lastRow / $chunkSize)
    Write-Host "Will create $totalFiles files" -ForegroundColor Cyan
    
    # Split the file
    for ($i = 1; $i -le $totalFiles; $i++) {
        Write-Host "Creating file $i of $totalFiles..." -ForegroundColor Green
        
        # Calculate row range for this chunk
        if ($i -eq 1) {
            $startRow = 1  # Include header
        } else {
            $startRow = ($i - 1) * $chunkSize + 1
        }
        $endRow = [math]::Min($i * $chunkSize, $lastRow)
        
        # Create new workbook
        $newWorkbook = $excel.Workbooks.Add()
        $newWorksheet = $newWorkbook.Sheets.Item(1)
        
        # Copy header row to all files
        if ($i -gt 1) {
            $headerRange = $worksheet.Range("A1", $worksheet.Cells(1, $lastCol))
            $headerRange.Copy()
            $newWorksheet.Range("A1").PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)
            
            # Copy data starting from row 2
            $dataRange = $worksheet.Range("A$startRow", $worksheet.Cells($endRow, $lastCol))
            $dataRange.Copy()
            $newWorksheet.Range("A2").PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)
        } else {
            # First file - copy everything including header
            $range = $worksheet.Range("A1", $worksheet.Cells($endRow, $lastCol))
            $range.Copy()
            $newWorksheet.Range("A1").PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)
        }
        
        # Save the new file
        $outputPath = "$outputFolder\TransactionData_Part$i.xlsx"
        $newWorkbook.SaveAs($outputPath)
        $newWorkbook.Close()
        
        Write-Host "Saved: TransactionData_Part$i.xlsx" -ForegroundColor Green
    }
    
    # Cleanup
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "SUCCESS! Files split completed!" -ForegroundColor Green
    Write-Host "Check folder: $outputFolder" -ForegroundColor Yellow
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Make sure Excel is closed
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
