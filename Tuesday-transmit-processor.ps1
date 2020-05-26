# on Tuesday transmit processor powershell script / File Converter
$USR_NAME         = $env:USERNAME
$USR_DL_FOLDER    = $env:USERPROFILE + "\Downloads"
$USR_DESKTOP      = $env:USERPROFILE + "\Desktop"
$CURRENT_DATE     = Get-Date -Format "yyyyMMdd"
#
$CSV_FILES = Get-ChildItem $USR_DL_FOLDER\$env:USERNAME*.csv -File
foreach ($item in $CSV_FILES) {
    $XLSX_HANDLE = New-Object -ComObject Excel.Application
    $XLSX_BOOK   = $XLSX_HANDLE.Workbooks.Open($item)
    $XLSX_HANDLE.DisplayAlerts = $false
    $XLSX_CONV =  ($item -replace ".csv","") # remove .csv from csv filename
    # can saved filename.xlsx not filename.csv.xlsx
    $XLSX_BOOK.SaveAs($XLSX_CONV, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
    $XLSX_HANDLE.Quit()
}
#
Move-Item -Path $USR_DL_FOLDER\$USR_NAME*.xlsx -Destination $USR_DESKTOP -Force
$XLSX_FILES = Get-ChildItem $USR_DESKTOP\$USR_NAME*.xlsx -File
foreach ($member in $XLSX_FILES) {
    # remove file path from $member for partial filename string replacement
    $RENAMED_XLSX = ((Get-ChildItem $member -Name) -replace ($USR_NAME), ($CURRENT_DATE))
    #write-host "NEW FILE NAME:　" $RENAMED_XLSX
    Rename-Item -Path $member -NewName $RENAMED_XLSX
}
