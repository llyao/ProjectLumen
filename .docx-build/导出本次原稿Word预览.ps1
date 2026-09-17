$ErrorActionPreference = 'Stop'
$项目目录 = Split-Path -Parent $PSScriptRoot
$清单 = Get-Content -LiteralPath (Join-Path $PSScriptRoot '本次原稿转Word清单.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$应用 = New-Object -ComObject Word.Application
$应用.Visible = $false
$应用.DisplayAlerts = 0
$应用.AutomationSecurity = 3
try {
    foreach ($文件名 in $清单.待转换) {
        $文档文件名 = [System.IO.Path]::ChangeExtension($文件名, '.docx')
        $卷号 = [regex]::Match($文件名, 'L\d{3}').Value
        $文档路径 = Join-Path $项目目录 ('朵朵\6年级\记录\英语\原稿\' + $文档文件名)
        $预览路径 = Join-Path $PSScriptRoot ($卷号 + '-新增原稿排版核对.pdf')
        $文档 = $应用.Documents.Open($文档路径, $false, $true)
        try {
            $null = $文档.Fields.Update()
            $文档.Repaginate()
            $文档.ExportAsFixedFormat($预览路径, 17)
            [PSCustomObject]@{卷号=$卷号; 页数=$文档.ComputeStatistics(2)} | ConvertTo-Json -Compress
        } finally {
            $文档.Close(0)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($文档)
        }
    }
} finally {
    $应用.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($应用)
}
