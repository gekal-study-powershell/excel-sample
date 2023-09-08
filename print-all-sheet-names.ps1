Param(
    [Parameter(Mandatory=$true, Position = 0)]
    [Alias('In')]
    [String]$ExcelPath
)

$excel = $null;
Function OpenExcel([String] $Path){

    try {
        $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application');
        foreach($workbook in $excel.workbooks ) {
            if ([String]::Equals($Path, $workbook.FullName)) {
                Write-Host ">>>>>>>>>> Opened Excel Reuse >>>>>>>>>>";
                return $workbook;
            }
        }
    } catch {
        Write-Host "Excel is not running.";
    }

    Write-Host ">>>>>>>>>> Open Excel >>>>>>>>>>";

    # Excelコンポーネント
    $global:excel = New-Object -ComObject "Excel.Application";

    $excel = $global:excel
    $excel.DisplayAlerts = $FALSE  # 警告を無視する。

    $workbook = $excel.Workbooks.Open($Path);;
    
    return $workbook;
}

# 存在チェック
$ExcelPath = $ExcelPath.Trim();
if (![System.IO.File]::Exists($ExcelPath)) {
    Write-Host "Excel is not exit.[${ExcelPath}]";
    exit;
}

try {
    $workBook = OpenExcel($ExcelPath);

    # >>>>>> Excelの操作 <<<<<<
    foreach ($sheet in $workbook.Worksheets) {
        Write-Host $sheet.name;
    }
    
} finally {

    if($null -ne $global:excel) {
        Write-Host ">>>>>>>>>> Close Excel <<<<<<<<<<";
        # Excelを終了する。
        $global:excel.Quit();
        $global:excel = $null;
    }

    [GC]::Collect();

    Write-Host "Finished."
}