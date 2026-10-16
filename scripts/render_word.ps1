param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory
)
# Windows 未提供 LibreOffice 時，以獨立的隱藏 Word instance 轉出 QA PDF。
# 只開啟明確指定的檔案，不操作使用者現有的 Word 視窗。
$ErrorActionPreference = 'Stop'
$sourcePath = (Resolve-Path -LiteralPath $InputPath).Path
$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
$destinationDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$wordApp = $null
try {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordApp.AutomationSecurity = 3
    $files = if (Test-Path -LiteralPath $sourcePath -PathType Container) {
        Get-ChildItem -LiteralPath $sourcePath -Filter '*.docx' -File
    } else { Get-Item -LiteralPath $sourcePath }
    foreach ($file in $files) {
        $document = $null
        try {
            $document = $wordApp.Documents.Open($file.FullName, $false, $true, $false)
            $pdfPath = Join-Path $destinationDirectory ($file.BaseName + '.pdf')
            $document.ExportAsFixedFormat($pdfPath, 17)
            Write-Output $pdfPath
        } finally {
            if ($null -ne $document) { $document.Close(0); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
        }
    }
} finally {
    if ($null -ne $wordApp) { $wordApp.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordApp) }
}
