
# Written by Sven De Preter 
# Convert XLS to CSV - PowerShell Version

Function ExcelToCSV($file){

    # Input : $file, a full path filename
    # Output : files with the same path and name as $file, with the worksheet name and .csv appended


    # Create new Excel Application object
    $ExcelObject = New-Object -ComObject Excel.Application
    $ExcelObject.Visible = $false
    $ExcelObject.DisplayAlerts=$false

    # Open the Excelfile 
    $workbook = $ExcelObject.Workbooks.Open($file.Fullname)
        
    # for each worksheet, save the file as a csv
    foreach ($worksheet in $workbook.Worksheets)
    {
        $worksheet.SaveAs(($file.Fullname -replace '.xlsx$', '') + "-" + ($worksheet.name) + ".csv",6)
    }

    # close and quit excel
    $excelObject.workbooks.close()
    $ExcelObject.quit();
}


# Files in the $Directory path will be processed
$Directory = "C:\\TEMP\\"

Foreach ($files in (Get-ChildItem -path $Directory -Filter "*.xlsx"))
{
    ExcelToCSV($files)
}
