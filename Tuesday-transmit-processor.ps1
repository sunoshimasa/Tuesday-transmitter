#
# on Tuesday transmit processor powershell script
#
# ダウンロードフォルダ内のCSV ファイル処理
#
$USR_NAME         = $env:USERNAME
$USR_DL_FOLDER    = $env:USERPROFILE + "\Downloads"
$USR_DESKTOP      = $env:USERPROFILE + "\Desktop"
$CURRENT_DATE     = Get-Date -Format "yyyyMMdd"
#
# ダウンロードフォルダ内の CSV ファイルをエクセル形式に変換
#
$CSV_FILES = Get-ChildItem $USR_DL_FOLDER\$env:USERNAME*.csv -File
foreach ($item in $CSV_FILES) {
    # エクセルのcomアプリケーションハンドルを取得
    $XLSX_HANDLE = New-Object -ComObject Excel.Application
    # CSV ファイルを開く
    $XLSX_BOOK   = $XLSX_HANDLE.Workbooks.Open($item)
    $XLSX_HANDLE.DisplayAlerts = $false
    # ファイル名から拡張子を削除して
    $XLSX_CONV =  ($item -replace ".csv","")
    # CSV ファイルを xlsxで保存　エクセル形式ファイル名の拡張子はxlsxが自動付与される
    $XLSX_BOOK.SaveAs($XLSX_CONV, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
    # comアプリケーションハンドル開放
    $XLSX_HANDLE.Quit()
}
    # エクセル形式ファイルをデスクトップに上書き移動
    Move-Item -Path $USR_DL_FOLDER\$USR_NAME*.xlsx -Destination $USR_DESKTOP -Force
    # 移動したデスクトップのエクセルファイルをリネーム
    $XLSX_FILES = Get-ChildItem $USR_DESKTOP\$USR_NAME*.xlsx -File
    foreach ($member in $XLSX_FILES) {
        # デスクトップのエクセルファイルのファイル名のユーザー名部分を日付に変更
        # ファイル名部分のみ文字列を置き換える
        $RENAMED_XLSX = ((Get-ChildItem $member -Name) -replace ($USR_NAME), ($CURRENT_DATE))
        # 変更後のファイル名にリネーム
        write-host "NEW FILE NAME:　" $RENAMED_XLSX
        Rename-Item -Path $member -NewName $RENAMED_XLSX
    }
