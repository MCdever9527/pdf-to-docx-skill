# render_docx.ps1 — 用本机 Word 把 .docx 渲染为 .pdf（验收闭环的渲染器）
#
# 为什么用 Word 而不是第三方转换器：
#   验收要回答的是「这份 docx 在 Word 里打开长什么样」，
#   渲染就必须是 Word 本体，否则分页/断行/字体回退的差异会掩盖真实问题。
#
# 用法:
#   pwsh -NoProfile -File render_docx.ps1 -InFile doc.docx -OutFile doc.pdf [-Quiet]
#
# 退出码: 0 成功 / 1 失败
# 注意: 参数名不能叫 $Input（PowerShell 自动变量），故用 InFile / OutFile

param(
    [Parameter(Mandatory = $true)][string]$InFile,
    [Parameter(Mandatory = $true)][string]$OutFile,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'

function Say($msg) { if (-not $Quiet) { Write-Output $msg } }

$inPath = [System.IO.Path]::GetFullPath($InFile)
$outPath = [System.IO.Path]::GetFullPath($OutFile)

if (-not (Test-Path -LiteralPath $inPath)) {
    Write-Output "ERROR 输入不存在: $inPath"
    exit 1
}

$outDir = [System.IO.Path]::GetDirectoryName($outPath)
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}
if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }

$word = $null
$doc = $null
$ok = $false
try {
    Say "STEP 启动 Word COM"
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.Options.WarnBeforeSavingPrintingSendingMarkup = $false } catch {}
    try { $word.Options.UpdateLinksAtOpen = $false } catch {}
    try { $word.AutomationSecurity = 3 } catch {}   # msoAutomationSecurityForceDisable

    Say "STEP 打开文档"
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ...)
    $doc = $word.Documents.Open($inPath, $false, $true, $false, "", "", $false)

    Say "STEP 导出 PDF"
    # 首选 SaveAs2：参数少、PowerShell 下 COM 绑定最稳
    # wdFormatPDF = 17
    try {
        $doc.SaveAs2($outPath, 17)
    }
    catch {
        Say "  SaveAs2 失败($($_.Exception.Message))，回退 SaveAs"
        $doc.SaveAs($outPath, 17)
    }

    if (Test-Path -LiteralPath $outPath) {
        $pages = 0
        try { $pages = $doc.ComputeStatistics(2) } catch {}   # wdStatisticPages
        $size = (Get-Item -LiteralPath $outPath).Length
        Say "RENDER_OK pages=$pages bytes=$size out=$outPath"
        $ok = $true
    }
    else {
        Say "RENDER_FAIL 未产生输出文件"
    }
}
catch {
    Say "RENDER_FAIL $($_.Exception.Message)"
}
finally {
    if ($doc) { try { $doc.Close(0) } catch {} }
    if ($word) { try { $word.Quit() } catch {} }
    if ($word) { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {} }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
if ($ok) { exit 0 } else { exit 1 }
