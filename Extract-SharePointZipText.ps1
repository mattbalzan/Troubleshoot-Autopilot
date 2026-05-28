<#
.SYNOPSIS
    Download a ZIP from SharePoint (using a pre-supplied cookie/token), extract it,
    pull text from PDF / DOCX / TXT / XLS(X) / PPT(X) / MSG (renamed to .html),
    and write everything into a single Word .docx with a Heading 1 per source file.

.PARAMETER ZipUrl
    Direct download URL to the .zip on SharePoint (e.g. the "download=1" share link).

.PARAMETER AuthCookie
    Cookie header value used to authenticate the download.
    For SharePoint Online this is typically the concatenation of the FedAuth and rtFa cookies,
    e.g.  "FedAuth=77u/PD94bWwg...; rtFa=KQ...="
    Grab them from a browser dev-tools session that is already signed in to the tenant.

.PARAMETER WorkRoot
    Folder where the zip is downloaded and extracted. Defaults to a timestamped folder under %TEMP%.

.PARAMETER OutputDocx
    Path of the combined Word document to produce. Defaults to <WorkRoot>\CombinedExtract.docx.

.EXAMPLE
    .\Extract-SharePointZipText.ps1 `
        -ZipUrl   'https://contoso.sharepoint.com/:u:/s/Team/EXXXXX?download=1' `
        -AuthCookie 'FedAuth=...; rtFa=...'

.NOTES
    Requires Windows + installed Microsoft Word, Excel and PowerPoint (uses COM automation).
    Word 2013+ is needed to open PDFs.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ZipUrl,
    [Parameter(Mandatory)] [string] $AuthCookie,
    [string] $WorkRoot   = (Join-Path $env:TEMP ("SPExtract_{0:yyyyMMdd_HHmmss}" -f (Get-Date))),
    [string] $OutputDocx
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 0. Prep working folder
# ---------------------------------------------------------------------------
New-Item -ItemType Directory -Force -Path $WorkRoot | Out-Null
$zipPath     = Join-Path $WorkRoot 'download.zip'
$extractRoot = Join-Path $WorkRoot 'extracted'
if (-not $OutputDocx) { $OutputDocx = Join-Path $WorkRoot 'CombinedExtract.docx' }

# ---------------------------------------------------------------------------
# 1. Download the zip with the supplied cookie
# ---------------------------------------------------------------------------
Write-Host "Downloading zip ..." -ForegroundColor Cyan
$headers = @{
    Cookie       = $AuthCookie
    'User-Agent' = 'Mozilla/5.0'
}
Invoke-WebRequest -Uri $ZipUrl -Headers $headers -OutFile $zipPath -UseBasicParsing

# ---------------------------------------------------------------------------
# 2. Extract
# ---------------------------------------------------------------------------
Write-Host "Extracting to $extractRoot ..." -ForegroundColor Cyan
if (Test-Path $extractRoot) { Remove-Item $extractRoot -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extractRoot -Force

# ---------------------------------------------------------------------------
# 3. Rename .msg -> .html (per request: treat MSG as HTML for text extraction)
# ---------------------------------------------------------------------------
Get-ChildItem -Path $extractRoot -Filter *.msg -Recurse -File | ForEach-Object {
    $target = [IO.Path]::ChangeExtension($_.FullName, '.html')
    if (Test-Path $target) { Remove-Item $target -Force }
    Rename-Item -LiteralPath $_.FullName -NewName ([IO.Path]::GetFileName($target))
}

# ---------------------------------------------------------------------------
# 4. Spin up Office COM
# ---------------------------------------------------------------------------
Write-Host "Starting Word ..." -ForegroundColor Cyan
$word                = New-Object -ComObject Word.Application
$word.Visible        = $false
$word.DisplayAlerts  = 0     # wdAlertsNone
$word.Options.ConfirmConversions = $false

$excel = $null
$ppt   = $null

# Style constants
$wdStyleHeading1 = -2
$wdStyleNormal   = -1
$wdFormatDocx    = 16
$wdCollapseEnd   = 0
$msoFalse        = 0
$msoTrue         = -1

# ---------------------------------------------------------------------------
# 5. Helpers
# ---------------------------------------------------------------------------
function Add-Heading {
    param($Doc, [string]$Text)
    $r = $Doc.Content
    $r.Collapse($wdCollapseEnd)
    $r.InsertParagraphAfter()
    $r.Collapse($wdCollapseEnd)
    $r.InsertAfter($Text)
    $r.Style = $wdStyleHeading1
    $r.Collapse($wdCollapseEnd)
    $r.InsertParagraphAfter()
}

function Add-Body {
    param($Doc, [string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { $Text = '(no extractable text)' }
    $r = $Doc.Content
    $r.Collapse($wdCollapseEnd)
    $r.Style = $wdStyleNormal
    $r.InsertAfter($Text.TrimEnd() + "`r`n`r`n")
}

function Get-WordableText {
    # Works for: .pdf, .doc, .docx, .rtf, .odt - anything Word can open.
    param([string]$Path, $WordApp)
    $doc = $WordApp.Documents.Open(
        $Path,      # FileName
        $false,     # ConfirmConversions
        $true,      # ReadOnly
        $false,     # AddToRecentFiles
        '',         # PasswordDocument
        '',         # PasswordTemplate
        $true,      # Revert
        '',         # WritePasswordDocument
        '',         # WritePasswordTemplate
        $null,      # Format
        $null,      # Encoding
        $false      # Visible
    )
    try   { return $doc.Content.Text }
    finally { $doc.Close($false) }
}

function Get-ExcelText {
    param([string]$Path)
    if (-not $script:excel) {
        $script:excel = New-Object -ComObject Excel.Application
        $script:excel.Visible       = $false
        $script:excel.DisplayAlerts = $false
    }
    $wb = $script:excel.Workbooks.Open($Path, 0, $true)   # ReadOnly
    try {
        $sb = [System.Text.StringBuilder]::new()
        foreach ($sheet in $wb.Worksheets) {
            [void]$sb.AppendLine("--- Sheet: $($sheet.Name) ---")
            $used = $sheet.UsedRange
            $vals = $used.Value2
            if ($null -eq $vals) { continue }
            if ($vals -is [object[,]]) {
                $rows = $vals.GetLength(0); $cols = $vals.GetLength(1)
                for ($r = 1; $r -le $rows; $r++) {
                    $cells = for ($c = 1; $c -le $cols; $c++) { "$($vals[$r,$c])" }
                    [void]$sb.AppendLine(($cells -join "`t"))
                }
            } else {
                [void]$sb.AppendLine("$vals")
            }
        }
        return $sb.ToString()
    } finally { $wb.Close($false) }
}

function Get-PowerPointText {
    param([string]$Path)
    if (-not $script:ppt) {
        $script:ppt = New-Object -ComObject PowerPoint.Application
    }
    # PowerPoint Open: FileName, ReadOnly, Untitled, WithWindow
    $pres = $script:ppt.Presentations.Open($Path, $msoTrue, $msoFalse, $msoFalse)
    try {
        $sb = [System.Text.StringBuilder]::new()
        foreach ($slide in $pres.Slides) {
            [void]$sb.AppendLine("--- Slide $($slide.SlideNumber) ---")
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    [void]$sb.AppendLine($shape.TextFrame.TextRange.Text)
                }
            }
            # Speaker notes
            if ($slide.NotesPage.Shapes.Count -gt 1) {
                $notes = $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text
                if ($notes) { [void]$sb.AppendLine("[notes] $notes") }
            }
        }
        return $sb.ToString()
    } finally { $pres.Close() }
}

function Get-HtmlText {
    param([string]$Path)
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    # drop script/style blocks first
    $clean = [regex]::Replace($raw, '(?is)<(script|style)[^>]*>.*?</\1>', '')
    # drop all remaining tags
    $clean = [regex]::Replace($clean, '<[^>]+>', ' ')
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    if ('System.Web.HttpUtility' -as [type]) {
        $clean = [System.Web.HttpUtility]::HtmlDecode($clean)
    }
    # collapse runs of whitespace
    return ([regex]::Replace($clean, '\s+', ' ')).Trim()
}

function Get-PlainText {
    param([string]$Path)
    return Get-Content -LiteralPath $Path -Raw -Encoding UTF8
}

# ---------------------------------------------------------------------------
# 6. Create output doc & walk extracted files
# ---------------------------------------------------------------------------
Write-Host "Building combined document ..." -ForegroundColor Cyan
$outDoc = $word.Documents.Add()

$files = Get-ChildItem -Path $extractRoot -File -Recurse | Sort-Object FullName
$i = 0
foreach ($f in $files) {
    $i++
    $rel = $f.FullName.Substring($extractRoot.Length).TrimStart('\','/')
    Write-Host ("[{0}/{1}] {2}" -f $i, $files.Count, $rel)

    $text = $null
    try {
        switch -Regex ($f.Extension.ToLowerInvariant()) {
            '^\.(pdf|docx?|rtf|odt)$'           { $text = Get-WordableText -Path $f.FullName -WordApp $word }
            '^\.(xlsx?|xlsm|xlsb)$'             { $text = Get-ExcelText    -Path $f.FullName }
            '^\.csv$'                           { $text = Get-PlainText    -Path $f.FullName }
            '^\.(pptx?|pptm)$'                  { $text = Get-PowerPointText -Path $f.FullName }
            '^\.html?$'                         { $text = Get-HtmlText     -Path $f.FullName }
            '^\.(txt|log|xml|json|ini|reg|md)$' { $text = Get-PlainText    -Path $f.FullName }
            default { $text = "(skipped: unsupported extension '$($f.Extension)')" }
        }
    } catch {
        $text = "(error extracting: $($_.Exception.Message))"
    }

    Add-Heading -Doc $outDoc -Text $rel
    Add-Body    -Doc $outDoc -Text $text
}

# ---------------------------------------------------------------------------
# 7. Save & cleanup
# ---------------------------------------------------------------------------
Write-Host "Saving $OutputDocx ..." -ForegroundColor Cyan
$outDoc.SaveAs([ref]$OutputDocx, [ref]$wdFormatDocx)
$outDoc.Close($false)

$word.Quit()
if ($excel) { $excel.Quit() }
if ($ppt)   { $ppt.Quit()   }

foreach ($com in @($word, $excel, $ppt)) {
    if ($com) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($com) }
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host "Done." -ForegroundColor Green
Write-Host "  Source zip : $zipPath"
Write-Host "  Extracted  : $extractRoot"
Write-Host "  Combined   : $OutputDocx"
