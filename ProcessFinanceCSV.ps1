$SourceDirectory = "D:\CustomerAccount\"
$OutputExcelFile = "D:\Output.xlsx"

Write-Host "INFO | Finance tool started"

# Remove OutputExcelFile if Exists
if (Test-Path $OutputExcelFile)
{
    try
    {
        Remove-Item $OutputExcelFile -ErrorAction Stop
        Write-Host "INFO | Removed existing output file: $OutputExcelFile"
    }
    catch
    {
        Write-Host "ERROR | Cannot remove existing output file: $OutputExcelFile"
        Read-Host -Prompt "INFO | Press Enter to exit"
        exit -1
    }
}

# Instantiate Excel Object
$excel = New-Object -ComObject Excel.Application
# Set Excel View
$excel.Visible = $false
# Set Display Alert
$excel.DisplayAlerts = $false
# Create New Excel Workbook
$wb = $excel.Workbooks.Add()

# Loop Through Each CSV in SourceDirectory
Get-ChildItem $SourceDirectory\*.csv | ForEach-Object {
    # Skip Empty CSV Files
    if ((Import-Csv $_.FullName).Length -gt 0) {
        Write-Host "INFO | Processing CSV file: $_.Name"
        # Open CSV File
        $csvBook = $excel.Workbooks.Open($_.FullName)
        # Set Sheet Index to Interact
        $csvSheet = $csvBook.Worksheets.Item(1)
        # Count All Rows That Has Value
        $rowCount = $csvSheet.UsedRange.rows.count

        # Identify Excel Workbook Location of Last Sheet
        $lastSheet = $wb.Worksheets | Select -Last 1
        # Create New Sheet After Last Sheet That Exists
        $xlsSheet = $wb.Worksheets.Add($lastSheet)
        # Set New Sheet Name to Date
        $xlsSheet.Name = $_.Name.Split('.')[0].Split('_')[1]

        # Copy Column D to Clipboard
        $csvSheet.Activate() | Out-Null
        $range = $csvSheet.Range("D2:D$rowCount")
        $range.Copy() | Out-Null

        # Paste Column D to A and Autofit Width
        $range = $xlsSheet.Range("A1:A$rowCount")
        $xlsSheet.Paste($range) | Out-Null
        $xlsSheet.UsedRange.Columns.Autofit() | Out-Null
        # Set Column to Number Format
        $xlsSheet.Range("A2:A$rowCount").NumberFormat = "0.0000"
        [System.Windows.Forms.Clipboard]::Clear()

        # Copy Column T to Clipboard
        $csvSheet.Activate() | Out-Null
        $range = $csvSheet.Range("T2:T$rowCount")
        $range.Copy() | Out-Null

        # Paste Column T to B and Autofit Width
        $range = $xlsSheet.Range("B1:B$rowCount")
        $xlsSheet.Paste($range) | Out-Null
        $xlsSheet.UsedRange.Columns.Autofit() | Out-Null
        [System.Windows.Forms.Clipboard]::Clear()

        # Copy Column CP to Clipboard
        $csvSheet.Activate() | Out-Null
        $range = $csvSheet.Range("CP2:CP$rowCount")
        $range.Copy() | Out-Null

        # Paste Column C to C and Autofit Width
        $range = $xlsSheet.Range("C1:C$rowCount")
        $xlsSheet.Paste($range) | Out-Null
        $xlsSheet.UsedRange.Columns.Autofit() | Out-Null
        [System.Windows.Forms.Clipboard]::Clear()

        # Close CSV
        $csvBook.Close($false) | Out-Null
    }
}

# Remove Sheet1
$wb.Sheets.Item('sheet1').Delete()
# Save Created Excel File
$wb.SaveAs($OutputExcelFile)
# Close Workbook
$wb.Close()
# Exit Excel Object
$excel.Quit()
Write-Host "INFO | Excel file generated: $OutputExcelFile"

Read-Host -Prompt "INFO | Process Completed. Press Enter to continue"
