$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false
$wordApp.DisplayAlerts = 0
$wordApp.AutomationSecurity = 3
try {
    $docPath = Join-Path $projectRoot '朵朵\6年级\记录\英语\错题\2026-09-16-朵朵英语错题集-L008至L010.docx'
    $pdfPath = Join-Path $PSScriptRoot '朵朵英语错题集-排版核对.pdf'
    $wordDoc = $wordApp.Documents.Open($docPath, $false, $true)
    try {
        $null = $wordDoc.Fields.Update()
        $wordDoc.Repaginate()
        $wordDoc.ExportAsFixedFormat($pdfPath, 17)
        [PSCustomObject]@{页数=$wordDoc.ComputeStatistics(2);节数=$wordDoc.Sections.Count} | ConvertTo-Json -Compress
    } finally { $wordDoc.Close(0); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordDoc) }
} finally { $wordApp.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) }
