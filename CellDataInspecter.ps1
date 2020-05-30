#
# File Data Inspecter
#
$USR_NAME         = $env:USERNAME
$USR_DL_FOLDER    = $env:USERPROFILE + "\Downloads"
$USR_DESKTOP      = $env:USERPROFILE + "\Desktop"
$CURRENT_DATE     = Get-Date -Format "yyyyMMdd"
#
# Inspect Excel Sheet Cell Value
#
$INSPECT_COLUMN = 1
$INSPECT_ROW    = 1
$INSPECT_CellData = "cell value"
$INSPECT_SHEET  = 1
#
$INSPECT_Xlsx = Get-ChildItem -Path $USR_DESKTOP -Name $CURRENT_DATE*.xlsx -File
write-host $INSPECT_Xlsx
foreach ($item in $INSPECT_Xlsx) {
    # Create COM Application Handle
    $XLSX_HANDLE = New-Object -ComObject Excel.Application
    # Open xlsx File
    $XLSX_BOOK   = $XLSX_HANDLE.Workbooks.Open($USR_DESKTOP+ "\\" +$item)
    $XLSX_HANDLE.DisplayAlerts = $false
    $XLSX_SHEET  = $XLSX_BOOK.worksheets.item($INSPECT_SHEET).Name
    write-host $XLSX_SHEET
    # Get Cell Value
    $CELL_VALUE = $XLSX_BOOK.worksheets.item($INSPECT_SHEET ).Cells($INSPECT_ROW ,$INSPECT_COLUMN).Text
    write-host $CELL_VALUE" at "$INSPECT_COLUMN","$INSPECT_ROW 
    # check cell value at $INSPECT_COLUMN, $INSPECT_ROW
    if ($CELL_VALUE -eq $INSPECT_CellData) {
        write-host ">>>>> "Specified Inspect condition match!!
        # Change File Extensiton to xlsxChange File Extensiton to xlsx
        # Change File Name to Another for Save
        $XLSX_CONV =  $USR_DESKTOP + "\" + $CURRENT_DATE + "-" + $CELL_VALUE+($item -replace ".csv","")
        write-host $XLSX_CONV
        # Save xlsx Files
        $XLSX_BOOK.SaveAs($XLSX_CONV, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
    } else {
        write-host no match...
    }
# Releas COM Application Handle
$XLSX_BOOK.close()
$XLSX_HANDLE.Quit()
write-host process finished
}
