$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$wordApp = $null
try {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordApp.AutomationSecurity = 3
    foreach ($paperCode in @('L008', 'L009', 'L010')) {
        $sourceDoc = Join-Path $projectRoot "朵朵\6年级\记录\英语\原稿\2026-09-16-${paperCode}英语试题.docx"
        $previewPdf = Join-Path $PSScriptRoot "${paperCode}-排版核对.pdf"
        $wordDoc = $wordApp.Documents.Open($sourceDoc, $false, $true)
        try {
            $null = $wordDoc.Fields.Update()
            $wordDoc.Repaginate()
            $pageCount = $wordDoc.ComputeStatistics(2)
            $wordDoc.ExportAsFixedFormat($previewPdf, 17)
            [PSCustomObject]@{试卷=$paperCode;页数=$pageCount} | ConvertTo-Json -Compress
        } finally {
            $wordDoc.Close(0)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordDoc)
        }
    }
} finally {
    if ($null -ne $wordApp) {
        $wordApp.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp)
    }
}
