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

    # Excel�R���|�[�l���g
    $global:excel = New-Object -ComObject "Excel.Application";

    $excel = $global:excel
    $excel.DisplayAlerts = $FALSE  # �x���𖳎�����B

    $workbook = $excel.Workbooks.Open($Path);;
    
    return $workbook;
}

# ���݃`�F�b�N
$ExcelPath = $ExcelPath.Trim();
if (![System.IO.File]::Exists($ExcelPath)) {
    Write-Host "Excel is not exit.[${ExcelPath}]";
    exit;
}

try {
    $workBook = OpenExcel($ExcelPath);

    # >>>>>> Excel�̑��� <<<<<<
    foreach ($sheet in $workbook.Worksheets) {
        Write-Host $sheet.name;
    }
    
} finally {

    if($null -ne $global:excel) {
        Write-Host ">>>>>>>>>> Close Excel <<<<<<<<<<";
        # Excel���I������B
        $global:excel.Quit();
        $global:excel = $null;
    }

    [GC]::Collect();

    Write-Host "Finished."
}